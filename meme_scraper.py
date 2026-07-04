import csv
import json
import logging
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urljoin
from urllib.robotparser import RobotFileParser
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup
from requests import Session
from requests.exceptions import RequestException, Timeout


# ============================================================
# KONFIGURACJA
# ============================================================

TARGET_URL = os.getenv(
    "TARGET_URL",
    "https://feargreedmeter.com/top-100-most-popular-meme-stocks-today",
)

DATA_DIR = os.getenv("DATA_DIR", "data")
LOCAL_TIMEZONE = os.getenv("LOCAL_TIMEZONE", "Europe/Warsaw")
REQUEST_TIMEOUT_SECONDS = int(os.getenv("REQUEST_TIMEOUT_SECONDS", "20"))

USER_AGENT = os.getenv(
    "USER_AGENT",
    "EducationalMemeStockCollector/1.0 "
    "(public GitHub Actions research project; low frequency; no bypassing)",
)

CHECK_ROBOTS_TXT = os.getenv("CHECK_ROBOTS_TXT", "true").lower() == "true"
STOP_IF_ROBOTS_UNAVAILABLE = os.getenv("STOP_IF_ROBOTS_UNAVAILABLE", "false").lower() == "true"

# robots.txt cache'ujemy, żeby nie pobierać go przy każdym uruchomieniu co 5 minut.
ROBOTS_CACHE_TTL_SECONDS = int(os.getenv("ROBOTS_CACHE_TTL_SECONDS", str(24 * 60 * 60)))

# Opcjonalne okno zbierania danych w czasie lokalnym.
# Przykład:
# COLLECTION_START_LOCAL="2026-07-05 00:00:00"
# COLLECTION_END_LOCAL="2026-07-19 00:00:00"
COLLECTION_START_LOCAL = os.getenv("COLLECTION_START_LOCAL", "").strip()
COLLECTION_END_LOCAL = os.getenv("COLLECTION_END_LOCAL", "").strip()

LAST_SUCCESS_JSON = os.path.join(DATA_DIR, "last_successful_snapshot.json")
ROBOTS_CACHE_JSON = os.path.join(DATA_DIR, "robots_cache.json")

CSV_HEADERS = [
    "data_pobrania",
    "kod",
    "nazwa",
    "upvotes",
    "mentions",
    "rank",
    "mention_change",
    "source_url",
]


# ============================================================
# LOGOWANIE
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

logger = logging.getLogger("meme-stock-collector")


# ============================================================
# WYJĄTKI
# ============================================================

class ScraperError(Exception):
    pass


class GracefulSkip(Exception):
    """Kończymy bez błędu, np. poza oknem dat albo przy 429."""


class BlockedScraperError(ScraperError):
    pass


class RateLimitedScraperError(ScraperError):
    pass


class FatalScraperError(ScraperError):
    pass


class TransientScraperError(ScraperError):
    pass


class ParseScraperError(ScraperError):
    pass


# ============================================================
# MODEL DANYCH
# ============================================================

@dataclass
class MemeStockRow:
    fetched_at_local: str
    fetched_at_utc: str
    rank: Optional[int]
    ticker: Optional[str]
    company_name: Optional[str]
    upvotes: Optional[int]
    mentions: Optional[int]
    mention_change: Optional[int]
    source_url: str
    raw_text: str

    def to_csv_dict(self) -> Dict[str, Any]:
        return {
            "data_pobrania": self.fetched_at_local,
            "kod": self.ticker,
            "nazwa": self.company_name,
            "upvotes": self.upvotes,
            "mentions": self.mentions,
            "rank": self.rank,
            "mention_change": self.mention_change,
            "source_url": self.source_url,
        }

    def to_json_dict(self) -> Dict[str, Any]:
        return {
            "fetched_at_local": self.fetched_at_local,
            "fetched_at_utc": self.fetched_at_utc,
            "rank": self.rank,
            "ticker": self.ticker,
            "company_name": self.company_name,
            "upvotes": self.upvotes,
            "mentions": self.mentions,
            "mention_change": self.mention_change,
            "source_url": self.source_url,
            "raw_text": self.raw_text,
        }


# ============================================================
# CZAS / PLIKI
# ============================================================

def local_tz() -> ZoneInfo:
    return ZoneInfo(LOCAL_TIMEZONE)


def get_now() -> Tuple[datetime, datetime]:
    now_utc = datetime.now(timezone.utc)
    now_local = now_utc.astimezone(local_tz())
    return now_utc, now_local


def iso(dt: datetime) -> str:
    return dt.isoformat(timespec="seconds")


def daily_csv_path(now_local: datetime) -> str:
    return os.path.join(DATA_DIR, f"meme_stocks_{now_local:%Y-%m-%d}.csv")


def parse_local_datetime(value: str) -> Optional[datetime]:
    value = value.strip()

    if not value:
        return None

    dt = datetime.fromisoformat(value.replace("T", " "))

    if dt.tzinfo is None:
        return dt.replace(tzinfo=local_tz())

    return dt.astimezone(local_tz())


def ensure_collection_window(now_local: datetime) -> None:
    start = parse_local_datetime(COLLECTION_START_LOCAL) if COLLECTION_START_LOCAL else None
    end = parse_local_datetime(COLLECTION_END_LOCAL) if COLLECTION_END_LOCAL else None

    if start and now_local < start:
        raise GracefulSkip(
            f"Poza oknem zbierania danych. Start: {iso(start)}, teraz: {iso(now_local)}"
        )

    if end and now_local > end:
        raise GracefulSkip(
            f"Poza oknem zbierania danych. Koniec: {iso(end)}, teraz: {iso(now_local)}"
        )


def read_json(path: str) -> Optional[Dict[str, Any]]:
    if not os.path.exists(path):
        return None

    try:
        with open(path, "r", encoding="utf-8") as file:
            return json.load(file)
    except (OSError, json.JSONDecodeError):
        return None


def atomic_write_json(payload: Dict[str, Any], path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)

    temp_path = f"{path}.tmp"

    with open(temp_path, "w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)

    os.replace(temp_path, path)


# ============================================================
# HTTP / ROBOTS.TXT
# ============================================================

def create_session() -> Session:
    session = requests.Session()

    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.8,pl;q=0.6",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
    })

    return session


def check_robots_txt_cached(session: Session) -> bool:
    os.makedirs(DATA_DIR, exist_ok=True)

    now_ts = int(time.time())
    cache = read_json(ROBOTS_CACHE_JSON)

    if cache and cache.get("target_url") == TARGET_URL:
        age = now_ts - int(cache.get("checked_at_unix", 0))

        if age < ROBOTS_CACHE_TTL_SECONDS:
            allowed = bool(cache.get("allowed", False))
            logger.info("Używam cache robots.txt: allowed=%s, wiek=%ss", allowed, age)
            return allowed

    robots_url = urljoin(TARGET_URL, "/robots.txt")
    logger.info("Sprawdzam robots.txt: %s", robots_url)

    try:
        response = session.get(robots_url, timeout=REQUEST_TIMEOUT_SECONDS)
    except RequestException as exc:
        if STOP_IF_ROBOTS_UNAVAILABLE:
            logger.error("Nie udało się pobrać robots.txt: %s. Przerywam.", exc)
            return False

        logger.warning("Nie udało się pobrać robots.txt: %s. Kontynuuję ostrożnie.", exc)
        return True

    if response.status_code == 404:
        if STOP_IF_ROBOTS_UNAVAILABLE:
            logger.error("robots.txt zwrócił 404. Przerywam.")
            return False

        logger.warning("robots.txt zwrócił 404. Kontynuuję ostrożnie.")
        return True

    if response.status_code != 200:
        if STOP_IF_ROBOTS_UNAVAILABLE:
            logger.error("robots.txt zwrócił HTTP %s. Przerywam.", response.status_code)
            return False

        logger.warning("robots.txt zwrócił HTTP %s. Kontynuuję ostrożnie.", response.status_code)
        return True

    parser = RobotFileParser()
    parser.set_url(robots_url)
    parser.parse(response.text.splitlines())

    allowed = parser.can_fetch(USER_AGENT, TARGET_URL)

    atomic_write_json({
        "target_url": TARGET_URL,
        "robots_url": robots_url,
        "checked_at_unix": now_ts,
        "checked_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "allowed": allowed,
    }, ROBOTS_CACHE_JSON)

    if allowed:
        logger.info("robots.txt pozwala na pobranie strony.")
    else:
        logger.error("robots.txt nie pozwala pobierać %s.", TARGET_URL)

    return allowed


def fetch_html(session: Session) -> str:
    logger.info("Pobieram stronę: %s", TARGET_URL)

    try:
        response = session.get(TARGET_URL, timeout=REQUEST_TIMEOUT_SECONDS)
    except Timeout as exc:
        raise TransientScraperError(f"Timeout połączenia: {exc}") from exc
    except RequestException as exc:
        raise TransientScraperError(f"Błąd połączenia: {exc}") from exc

    status = response.status_code

    if status == 200:
        return response.text

    if status == 403:
        raise BlockedScraperError(
            "HTTP 403: strona odmówiła dostępu. Nie ponawiam agresywnie."
        )

    if status == 404:
        raise FatalScraperError("HTTP 404: strona nie została znaleziona.")

    if status == 429:
        raise RateLimitedScraperError(
            "HTTP 429: limit zapytań. Nie ponawiam agresywnie."
        )

    if 500 <= status <= 599:
        raise TransientScraperError(f"HTTP {status}: błąd po stronie serwera.")

    if 400 <= status <= 499:
        raise FatalScraperError(f"HTTP {status}: błąd klienta.")

    raise TransientScraperError(f"Nieoczekiwany HTTP status: {status}")


# ============================================================
# PARSOWANIE
# ============================================================

def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def parse_int(value: Optional[str]) -> Optional[int]:
    if value is None:
        return None

    value = value.strip()
    value = value.replace(",", "")
    value = value.replace("$", "")
    value = value.replace("%", "")
    value = value.replace("−", "-")
    value = re.sub(r"\s+", "", value)

    if not value:
        return None

    try:
        return int(value)
    except ValueError:
        return None


def looks_like_rank_row(text: str) -> bool:
    if not text:
        return False

    return bool(
        re.match(
            r"^#\s*\d{1,3}\s*\.?\s+[A-Z][A-Z0-9.\-]{0,9}\b",
            text,
        )
    )


def pop_trailing_signed_int(tokens: List[str]) -> Tuple[Optional[int], List[str]]:
    if not tokens:
        return None, tokens

    # Format: + 575
    if len(tokens) >= 2 and tokens[-2] in {"+", "-"}:
        number = parse_int(tokens[-1])

        if number is not None:
            sign = 1 if tokens[-2] == "+" else -1
            return sign * abs(number), tokens[:-2]

    # Format: +575 albo -6
    last = tokens[-1].replace("−", "-")

    if re.match(r"^[+-]\d[\d,]*$", last):
        return parse_int(last), tokens[:-1]

    return None, tokens


def parse_ranked_text_row(
    text: str,
    fetched_at_local: str,
    fetched_at_utc: str,
) -> Optional[MemeStockRow]:
    """
    Parser dla formatu podobnego do:

        # 1 . MU Micron Technology 2,893 677 + 575

    Wynik:
        rank = 1
        ticker = MU
        company_name = Micron Technology
        upvotes = 2893
        mentions = 677
        mention_change = 575
    """
    match = re.match(
        r"^#\s*(?P<rank>\d{1,3})\s*\.?\s+"
        r"(?P<ticker>[A-Z][A-Z0-9.\-]{0,9})\s+"
        r"(?P<rest>.+)$",
        text,
    )

    if not match:
        return None

    rank = parse_int(match.group("rank"))
    ticker = match.group("ticker").strip().upper()
    tokens = clean_text(match.group("rest")).split()

    if len(tokens) < 3:
        company_name = " ".join(tokens) or None

        return MemeStockRow(
            fetched_at_local=fetched_at_local,
            fetched_at_utc=fetched_at_utc,
            rank=rank,
            ticker=ticker,
            company_name=company_name,
            upvotes=None,
            mentions=None,
            mention_change=None,
            source_url=TARGET_URL,
            raw_text=text,
        )

    mention_change, tokens = pop_trailing_signed_int(tokens)

    mentions = None
    upvotes = None

    if tokens:
        candidate = parse_int(tokens[-1])

        if candidate is not None:
            mentions = candidate
            tokens = tokens[:-1]

    if tokens:
        candidate = parse_int(tokens[-1])

        if candidate is not None:
            upvotes = candidate
            tokens = tokens[:-1]

    company_name = " ".join(tokens).strip() or None

    return MemeStockRow(
        fetched_at_local=fetched_at_local,
        fetched_at_utc=fetched_at_utc,
        rank=rank,
        ticker=ticker,
        company_name=company_name,
        upvotes=upvotes,
        mentions=mentions,
        mention_change=mention_change,
        source_url=TARGET_URL,
        raw_text=text,
    )


def deduplicate_rows(rows: List[MemeStockRow]) -> List[MemeStockRow]:
    seen = set()
    result: List[MemeStockRow] = []

    for row in rows:
        key = (row.rank, row.ticker)

        if key in seen:
            continue

        seen.add(key)
        result.append(row)

    result.sort(key=lambda row: row.rank if row.rank is not None else 9999)

    return result


def parse_html(
    html: str,
    fetched_at_local: str,
    fetched_at_utc: str,
) -> List[MemeStockRow]:
    soup = BeautifulSoup(html, "lxml")
    candidates: List[str] = []

    # Obecnie ranking jest dostępny jako tekst linków/kart.
    for element in soup.select("a"):
        text = clean_text(element.get_text(" ", strip=True))

        if looks_like_rank_row(text):
            candidates.append(text)

    # Fallback, gdyby linki zmieniły się na divy/karty.
    if not candidates:
        for selector in [
            "[class*='stock']",
            "[class*='rank']",
            "[class*='ticker']",
            "[class*='card']",
            "[class*='item']",
        ]:
            for element in soup.select(selector):
                text = clean_text(element.get_text(" ", strip=True))

                if looks_like_rank_row(text):
                    candidates.append(text)

    rows: List[MemeStockRow] = []

    for text in candidates:
        row = parse_ranked_text_row(text, fetched_at_local, fetched_at_utc)

        if row:
            rows.append(row)

    rows = deduplicate_rows(rows)

    if not rows:
        raise ParseScraperError(
            "Nie znaleziono rekordów rankingu. "
            "Strona mogła zmienić HTML albo zacząć ładować dane JavaScriptem."
        )

    return rows


# ============================================================
# ZAPIS DO CSV / JSON
# ============================================================

def append_rows_to_daily_csv(rows: List[MemeStockRow], path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)

    file_is_empty = not os.path.exists(path) or os.path.getsize(path) == 0

    with open(path, "a", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_HEADERS)

        if file_is_empty:
            writer.writeheader()

        for row in rows:
            writer.writerow(row.to_csv_dict())

    logger.info("Dopisano %s rekordów do CSV: %s", len(rows), path)


def save_last_successful_snapshot(rows: List[MemeStockRow]) -> None:
    atomic_write_json({
        "saved_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "source_url": TARGET_URL,
        "rows_count": len(rows),
        "rows": [row.to_json_dict() for row in rows],
    }, LAST_SUCCESS_JSON)

    logger.info("Zapisano ostatni poprawny snapshot: %s", LAST_SUCCESS_JSON)


# ============================================================
# MAIN — JEDNO POBRANIE I KONIEC
# ============================================================

def run_once() -> None:
    now_utc, now_local = get_now()

    ensure_collection_window(now_local)

    fetched_at_utc = iso(now_utc)
    fetched_at_local = iso(now_local)
    csv_path = daily_csv_path(now_local)

    logger.info("Start jednorazowego pobrania.")
    logger.info("Czas lokalny: %s", fetched_at_local)
    logger.info("Czas UTC: %s", fetched_at_utc)
    logger.info("Plik dzienny CSV: %s", csv_path)

    session = create_session()

    if CHECK_ROBOTS_TXT and not check_robots_txt_cached(session):
        raise GracefulSkip("robots.txt nie pozwala na pobranie strony.")

    html = fetch_html(session)
    rows = parse_html(html, fetched_at_local, fetched_at_utc)

    if len(rows) < 50:
        logger.warning(
            "Znaleziono tylko %s rekordów. Sprawdź, czy parser nadal pasuje.",
            len(rows),
        )

    append_rows_to_daily_csv(rows, csv_path)
    save_last_successful_snapshot(rows)

    logger.info("Pobranie zakończone sukcesem.")


def main() -> int:
    try:
        run_once()
        return 0

    except GracefulSkip as exc:
        logger.warning("Kończę bez błędu: %s", exc)
        return 0

    except (BlockedScraperError, RateLimitedScraperError) as exc:
        logger.warning("Kończę bez błędu, żeby nie ponawiać agresywnie: %s", exc)
        return 0

    except ScraperError as exc:
        logger.exception("Scraper zakończył się błędem: %s", exc)
        return 1

    except Exception as exc:
        logger.exception("Nieoczekiwany błąd: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
