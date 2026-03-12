"""
export_automaty.py – Miesięczne zestawienie TVM wg schematu 1-8.

Generuje raport P&L dla automatów biletowych TVM za wybrany miesiąc.
Domyślnie przetwarza poprzedni pełny miesiąc; można podać YYYY-MM ręcznie.

Uruchamianie:
    python export_automaty.py                    # poprzedni miesiąc (auto)
    python export_automaty.py --month 2026-02    # wybrany miesiąc
    python export_automaty.py --month 2026-02 --config config.local.yaml

Schemat kolumn wyjściowych (1-8):
    1-INFO            informacje opisowe o automacie (zmiana-tracked)
    2-OBRÓT           obrót miesięczny z bazy danych
    3-PROWIZJA        prowizja (ELAVON + interchange)
    4-PODATEK         podatek VAT
    5-KOSZTY          koszty zewnętrzne (czynsz, prąd, ELAVON, …)
    6-PO SUMIE        wynik netto po wszystkich składowych
    7-PODSUMOWANIE ROCZNE  agregacja roczna (YTD)
    8-UWAGI           uwagi manualne / automatyczne
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import smtplib
import sys
from calendar import monthrange
from datetime import date, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

import yaml

# Optional heavy dependencies – imported lazily so unit tests can run
# without a full environment.
try:
    import pandas as pd
    import openpyxl  # noqa: F401 – required by pandas Excel writer
    _PANDAS_AVAILABLE = True
except ImportError:
    _PANDAS_AVAILABLE = False

try:
    import urllib.parse
    import sqlalchemy
    _SQLALCHEMY_AVAILABLE = True
except ImportError:
    _SQLALCHEMY_AVAILABLE = False

try:
    import requests as _requests
    _REQUESTS_AVAILABLE = True
except ImportError:
    _REQUESTS_AVAILABLE = False


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

DEFAULT_CONFIG_PATH = Path(__file__).parent / "config.yaml"


def load_config(path: str | Path = DEFAULT_CONFIG_PATH) -> dict:
    """Load YAML configuration file."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Plik konfiguracyjny nie istnieje: {path}")
    with path.open(encoding="utf-8") as fh:
        cfg = yaml.safe_load(fh)
    return cfg or {}


# ---------------------------------------------------------------------------
# Reporting period helpers
# ---------------------------------------------------------------------------

def resolve_reporting_period(month_arg: str | None, cfg: dict) -> tuple[int, int]:
    """Return (year, month) for the reporting period.

    Priority:
        1. CLI argument --month YYYY-MM
        2. config.yaml reporting.default_period (if not "auto")
        3. Previous full calendar month
    """
    if month_arg:
        return _parse_yyyymm(month_arg)

    default = cfg.get("reporting", {}).get("default_period", "auto")
    if default and default != "auto":
        return _parse_yyyymm(str(default))

    # Previous full month
    today = date.today()
    if today.month == 1:
        return today.year - 1, 12
    return today.year, today.month - 1


def _parse_yyyymm(value: str) -> tuple[int, int]:
    try:
        dt = datetime.strptime(value.strip(), "%Y-%m")
        return dt.year, dt.month
    except ValueError as exc:
        raise ValueError(
            f"Nieprawidłowy format okresu: '{value}'. Oczekiwano YYYY-MM."
        ) from exc


def month_date_range(year: int, month: int) -> tuple[date, date]:
    """Return (first_day, last_day) for the given year/month."""
    last_day = monthrange(year, month)[1]
    return date(year, month, 1), date(year, month, last_day)


# ---------------------------------------------------------------------------
# Database connection
# ---------------------------------------------------------------------------

_SAFE_IDENTIFIER_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def _validate_sql_identifier(name: str) -> str:
    """Raise ValueError if *name* is not a safe SQL identifier (alphanumeric + _)."""
    if not _SAFE_IDENTIFIER_RE.match(name):
        raise ValueError(
            f"Nieprawidłowa nazwa tabeli/kolumny: {name!r}. "
            "Dozwolone znaki: litery, cyfry i podkreślenie."
        )
    return name


def build_connection_string(db_cfg: dict) -> str:
    """Build a SQLAlchemy connection string from config."""
    db_type = db_cfg.get("type", "mssql").lower()
    host = db_cfg.get("host", "localhost")
    port = db_cfg.get("port", "")
    name = db_cfg.get("name", "")
    user = db_cfg.get("user", "")
    # Password: config value OR environment variable DB_PASSWORD (env takes priority)
    password = os.environ.get("DB_PASSWORD") or db_cfg.get("password", "") or ""
    password_encoded = urllib.parse.quote_plus(str(password)) if password else ""

    port_str = f":{port}" if port else ""

    if db_type == "mssql":
        driver = "ODBC+Driver+17+for+SQL+Server"
        return (
            f"mssql+pyodbc://{user}:{password_encoded}@{host}{port_str}/{name}"
            f"?driver={driver}"
        )
    if db_type == "mysql":
        return f"mysql+pymysql://{user}:{password_encoded}@{host}{port_str}/{name}"
    if db_type == "postgresql":
        return f"postgresql+psycopg2://{user}:{password_encoded}@{host}{port_str}/{name}"
    if db_type == "sqlite":
        return f"sqlite:///{name}"

    raise ValueError(f"Nieznany typ bazy danych: {db_type}")


def get_db_engine(cfg: dict):
    """Create a SQLAlchemy engine from config. Raises if not available."""
    if not _SQLALCHEMY_AVAILABLE:
        raise RuntimeError(
            "Pakiet sqlalchemy nie jest zainstalowany. "
            "Uruchom: pip install sqlalchemy"
        )
    conn_str = build_connection_string(cfg["database"])
    return sqlalchemy.create_engine(conn_str)


# ---------------------------------------------------------------------------
# 1-INFO – machine information with change tracking
# ---------------------------------------------------------------------------

INFO_FIELDS = [
    "nr_automatu", "lokalizacja", "przewoznik", "operator",
    "segment", "status", "data_instalacji",
]


def fetch_info_from_db(engine, year: int, month: int) -> list[dict]:
    """Fetch 1-INFO records from the database."""
    table = _validate_sql_identifier("automaty")
    cols = ", ".join(_validate_sql_identifier(f) for f in INFO_FIELDS)
    query = f"SELECT {cols} FROM {table} WHERE status != 'USUNIETY'"
    if not _PANDAS_AVAILABLE:
        raise RuntimeError("Pakiet pandas jest wymagany do pobierania danych.")
    df = pd.read_sql(query, engine)
    return df.to_dict(orient="records")


def _snapshot_path(snapshots_dir: Path, year: int, month: int) -> Path:
    return snapshots_dir / f"info_snapshot_{year}_{month:02d}.json"


def load_snapshot(snapshots_dir: Path, year: int, month: int) -> dict[str, dict]:
    """Load a previous INFO snapshot keyed by nr_automatu."""
    path = _snapshot_path(snapshots_dir, year, month)
    if not path.exists():
        return {}
    with path.open(encoding="utf-8") as fh:
        data = json.load(fh)
    return {r["nr_automatu"]: r for r in data}


def save_snapshot(snapshots_dir: Path, year: int, month: int, records: list[dict]) -> None:
    """Persist INFO records as a snapshot for the given month."""
    snapshots_dir.mkdir(parents=True, exist_ok=True)
    path = _snapshot_path(snapshots_dir, year, month)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(records, fh, ensure_ascii=False, default=str, indent=2)
    log.info("Snapshot INFO zapisany: %s", path)


def detect_info_changes(
    current: list[dict],
    previous_snapshot: dict[str, dict],
) -> dict[str, list[str]]:
    """Return dict of nr_automatu → list of changed field descriptions."""
    changes: dict[str, list[str]] = {}
    for rec in current:
        key = rec["nr_automatu"]
        prev = previous_snapshot.get(key)
        if prev is None:
            changes[key] = ["NOWY AUTOMAT"]
            continue
        diffs = []
        for field in INFO_FIELDS:
            if str(rec.get(field)) != str(prev.get(field)):
                diffs.append(f"{field}: {prev.get(field)} → {rec.get(field)}")
        if diffs:
            changes[key] = diffs
    return changes


def process_info(
    engine,
    year: int,
    month: int,
    snapshots_dir: Path,
) -> tuple[list[dict], dict[str, list[str]]]:
    """Fetch current INFO, compare with previous snapshot, save new snapshot.

    Returns:
        (info_records, changes_dict)
    """
    # Previous month
    if month == 1:
        prev_year, prev_month = year - 1, 12
    else:
        prev_year, prev_month = year, month - 1

    info_records = fetch_info_from_db(engine, year, month)
    prev_snapshot = load_snapshot(snapshots_dir, prev_year, prev_month)
    changes = detect_info_changes(info_records, prev_snapshot)
    save_snapshot(snapshots_dir, year, month, info_records)

    if changes:
        log.info(
            "Wykryto zmiany INFO dla %d automatów: %s",
            len(changes),
            list(changes.keys()),
        )
    else:
        log.info("Brak zmian w danych INFO względem poprzedniego miesiąca.")

    return info_records, changes


# ---------------------------------------------------------------------------
# 2-OBRÓT – monthly turnover from DB
# ---------------------------------------------------------------------------

def fetch_obrot_from_db(engine, year: int, month: int, table: str) -> list[dict]:
    """Fetch turnover records for the given month."""
    table = _validate_sql_identifier(table)
    query = f"""
        SELECT
            nr_automatu,
            SUM(kwota_brutto) AS obrot_brutto,
            COUNT(*) AS liczba_transakcji
        FROM {table}
        WHERE YEAR(data_transakcji) = :year
          AND MONTH(data_transakcji) = :month
        GROUP BY nr_automatu
    """
    if not _PANDAS_AVAILABLE:
        raise RuntimeError("Pakiet pandas jest wymagany.")
    df = pd.read_sql(query, engine, params={"year": int(year), "month": int(month)})
    df["srednia_transakcja"] = df.apply(
        lambda row: (
            round(row["obrot_brutto"] / row["liczba_transakcji"], 2)
            if row["liczba_transakcji"] > 0
            else 0.0
        ),
        axis=1,
    )
    return df.to_dict(orient="records")


# ---------------------------------------------------------------------------
# 3-PROWIZJA – commission calculation and DB fetch
# ---------------------------------------------------------------------------

COMMISSION_FIELDS = [
    "prowizja_elavon",
    "prowizja_interchange",
    "prowizja_total",
    "stawka_elavon_pct",
    "stawka_interchange_pct",
]


def calculate_commission(
    obrot_brutto: float,
    przewoznik: str,
    commission_cfg: dict,
) -> dict:
    """Calculate ELAVON and interchange commissions from config rates."""
    elavon_rate = float(commission_cfg.get("elavon_default", 0.0125))

    interchange_rates = commission_cfg.get("interchange_per_carrier", {})
    interchange_rate = float(
        interchange_rates.get(przewoznik)
        or interchange_rates.get("default", 0.01)
    )

    prowizja_elavon = round(obrot_brutto * elavon_rate, 2)
    prowizja_interchange = round(obrot_brutto * interchange_rate, 2)
    prowizja_total = round(prowizja_elavon + prowizja_interchange, 2)

    return {
        "prowizja_elavon": prowizja_elavon,
        "prowizja_interchange": prowizja_interchange,
        "prowizja_total": prowizja_total,
        "stawka_elavon_pct": round(elavon_rate * 100, 4),
        "stawka_interchange_pct": round(interchange_rate * 100, 4),
    }


def fetch_prowizja_from_db(
    engine: Any, year: int, month: int, table: str
) -> dict[str, dict]:
    """Fetch commission records for the given month from the database.

    Expected table columns: nr_automatu, rok, miesiac,
    prowizja_elavon, prowizja_interchange, prowizja_total,
    stawka_elavon_pct, stawka_interchange_pct.

    Returns a dict keyed by nr_automatu.
    """
    table = _validate_sql_identifier(table)
    cols = ", ".join(_validate_sql_identifier(f) for f in COMMISSION_FIELDS)
    query = f"""
        SELECT nr_automatu, {cols}
        FROM {table}
        WHERE rok = :year AND miesiac = :month
    """
    if not _PANDAS_AVAILABLE:
        raise RuntimeError("Pakiet pandas jest wymagany.")
    df = pd.read_sql(query, engine, params={"year": int(year), "month": int(month)})
    result: dict[str, dict] = {}
    for _, row in df.iterrows():
        key = str(row["nr_automatu"])
        result[key] = {
            field: float(row.get(field, 0) or 0) for field in COMMISSION_FIELDS
        }
    return result


def load_prowizja(
    cfg: dict,
    engine: Any,
    year: int,
    month: int,
    info_map: dict[str, dict],
    obrot_map: dict[str, dict],
) -> dict[str, dict]:
    """Commission source adapter: 'calculated' (config rates) | 'database'.

    Returns a dict keyed by nr_automatu with commission fields.
    When source is 'calculated', commissions are derived from config rates
    and the turnover already loaded for each automat.
    When source is 'database', values are fetched directly from a DB table,
    mirroring how turnover is fetched from the transakcje table.
    """
    commission_cfg = cfg.get("commission_rates", {})
    source = commission_cfg.get("source", "calculated").lower()

    if source == "database":
        table = cfg["database"]["tables"]["prowizje"]
        db_records = fetch_prowizja_from_db(engine, year, month, table)
        log.info(
            "Pobrano dane prowizji z bazy danych dla %d automatów.", len(db_records)
        )
        return db_records

    # "calculated" – derive from config rates and turnover already loaded
    result: dict[str, dict] = {}
    for nr, info in info_map.items():
        obrot_brutto = float(obrot_map.get(nr, {}).get("obrot_brutto", 0))
        result[nr] = calculate_commission(
            obrot_brutto, info.get("przewoznik", ""), commission_cfg
        )
    log.info("Prowizja obliczona ze stawek konfiguracyjnych dla %d automatów.", len(result))
    return result


# ---------------------------------------------------------------------------
# 4-PODATEK – VAT calculation
# ---------------------------------------------------------------------------

def calculate_tax(obrot_brutto: float, tax_cfg: dict) -> dict:
    """Calculate VAT and net turnover."""
    vat_rate = float(tax_cfg.get("vat_rate", 0.23))
    method = tax_cfg.get("method", "from_gross")

    if method == "from_gross":
        # VAT zawarty w cenie (metoda „od stu")
        vat_nalezny = round(obrot_brutto * vat_rate / (1 + vat_rate), 2)
    else:
        # VAT doliczany do ceny (metoda „do stu")
        vat_nalezny = round(obrot_brutto * vat_rate, 2)

    obrot_netto = round(obrot_brutto - vat_nalezny, 2)

    return {
        "obrot_netto": obrot_netto,
        "vat_nalezny": vat_nalezny,
        "stawka_vat_pct": round(vat_rate * 100, 2),
    }


# ---------------------------------------------------------------------------
# 5-KOSZTY – cost aggregation (CSV / Excel / DB adapter)
# ---------------------------------------------------------------------------

COST_FIELDS = [
    "koszt_czynsz",
    "koszt_prad",
    "koszt_elavon",
    "koszt_poczta",
    "koszt_amortyzacja",
    "koszt_serwis",
    "koszt_transmisja",
    "koszt_pozostale",
]


def load_costs_from_csv(
    csv_path: Path, year: int, month: int
) -> dict[str, dict]:
    """Load cost records from a CSV file for the given year/month."""
    if not csv_path.exists():
        log.warning("Plik kosztów CSV nie istnieje: %s", csv_path)
        return {}
    costs: dict[str, dict] = {}
    with csv_path.open(encoding="utf-8", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            try:
                if int(row["rok"]) == year and int(row["miesiac"]) == month:
                    key = row["nr_automatu"]
                    costs[key] = {
                        field: float(row.get(field, 0) or 0)
                        for field in COST_FIELDS
                    }
                    costs[key]["uwagi_koszty"] = row.get("uwagi_koszty", "")
            except (KeyError, ValueError) as exc:
                log.warning("Błąd parsowania wiersza kosztów CSV: %s – %s", row, exc)
    return costs


def load_costs_from_excel(
    excel_path: Path, sheet: str, year: int, month: int
) -> dict[str, dict]:
    """Load cost records from an Excel file for the given year/month."""
    if not _PANDAS_AVAILABLE:
        raise RuntimeError("Pakiet pandas jest wymagany do wczytania Excel.")
    if not excel_path.exists():
        log.warning("Plik kosztów Excel nie istnieje: %s", excel_path)
        return {}
    df = pd.read_excel(excel_path, sheet_name=sheet)
    df_filtered = df[(df["rok"] == year) & (df["miesiac"] == month)]
    costs: dict[str, dict] = {}
    for _, row in df_filtered.iterrows():
        key = str(row["nr_automatu"])
        costs[key] = {
            field: float(row.get(field, 0) or 0) for field in COST_FIELDS
        }
        costs[key]["uwagi_koszty"] = str(row.get("uwagi_koszty", "") or "")
    return costs


def load_costs_from_db(engine, year: int, month: int, table: str) -> dict[str, dict]:
    """Load cost records from the database."""
    table = _validate_sql_identifier(table)
    cols = ", ".join(_validate_sql_identifier(f) for f in COST_FIELDS)
    query = f"""
        SELECT nr_automatu, {cols}, uwagi_koszty
        FROM {table}
        WHERE rok = :year AND miesiac = :month
    """
    if not _PANDAS_AVAILABLE:
        raise RuntimeError("Pakiet pandas jest wymagany.")
    df = pd.read_sql(query, engine, params={"year": int(year), "month": int(month)})
    costs: dict[str, dict] = {}
    for _, row in df.iterrows():
        key = str(row["nr_automatu"])
        costs[key] = {field: float(row.get(field, 0) or 0) for field in COST_FIELDS}
        costs[key]["uwagi_koszty"] = str(row.get("uwagi_koszty", "") or "")
    return costs


def load_costs(
    cfg: dict, engine, year: int, month: int
) -> dict[str, dict]:
    """Cost source adapter: CSV | Excel | database."""
    costs_cfg = cfg.get("costs", {})
    source = costs_cfg.get("source", "csv").lower()

    if source == "csv":
        csv_path = Path(costs_cfg.get("csv_path", "kas.csv"))
        return load_costs_from_csv(csv_path, year, month)

    if source == "excel":
        excel_path = Path(costs_cfg.get("excel_path", "kas.xlsx"))
        sheet = costs_cfg.get("excel_sheet", "Koszty")
        return load_costs_from_excel(excel_path, sheet, year, month)

    if source == "database":
        table = cfg["database"]["tables"]["koszty"]
        return load_costs_from_db(engine, year, month, table)

    raise ValueError(f"Nieznane źródło kosztów: {source!r}. Oczekiwano: csv | excel | database")


def aggregate_costs(cost_record: dict) -> float:
    """Sum all cost fields."""
    return round(sum(float(cost_record.get(f, 0)) for f in COST_FIELDS), 2)


# ---------------------------------------------------------------------------
# 6-PO SUMIE – net result
# ---------------------------------------------------------------------------

def calculate_wynik(
    obrot_netto: float,
    obrot_brutto: float,
    prowizja_total: float,
    koszty_total: float,
    alerts_cfg: dict,
) -> dict:
    """Calculate net result (wynik netto) and margin."""
    wynik_netto = round(obrot_netto - prowizja_total - koszty_total, 2)
    wynik_brutto = round(obrot_brutto - prowizja_total - koszty_total, 2)
    marza_netto_pct = (
        round(wynik_netto / obrot_netto * 100, 2) if obrot_netto else 0.0
    )

    alert = ""
    threshold = float(alerts_cfg.get("wynik_netto_min", -500))
    if wynik_netto < threshold:
        alert = f"ALERT: wynik netto {wynik_netto:.2f} poniżej progu {threshold:.2f}"

    return {
        "wynik_netto": wynik_netto,
        "wynik_brutto": wynik_brutto,
        "marza_netto_pct": marza_netto_pct,
        "alert_wynik": alert,
    }


# ---------------------------------------------------------------------------
# 7-PODSUMOWANIE ROCZNE – annual summary (YTD)
# ---------------------------------------------------------------------------

def build_annual_summary(monthly_rows: list[dict], year: int) -> dict:
    """Aggregate monthly rows for the given year (YTD)."""
    rows_year = [r for r in monthly_rows if r.get("rok") == year]
    return {
        "rok": year,
        "obrot_brutto_ytd": round(sum(r.get("obrot_brutto", 0) for r in rows_year), 2),
        "prowizja_ytd": round(sum(r.get("prowizja_total", 0) for r in rows_year), 2),
        "koszty_ytd": round(sum(r.get("koszty_total", 0) for r in rows_year), 2),
        "wynik_netto_ytd": round(sum(r.get("wynik_netto", 0) for r in rows_year), 2),
        "liczba_miesiecy": len({r.get("miesiac") for r in rows_year}),
    }


def build_annual_summary_per_automat(
    monthly_rows: list[dict], year: int
) -> dict[str, dict]:
    """Build YTD summary per automat."""
    automaty: dict[str, list[dict]] = {}
    for row in monthly_rows:
        if row.get("rok") == year:
            key = row["nr_automatu"]
            automaty.setdefault(key, []).append(row)

    summaries: dict[str, dict] = {}
    for nr, rows in automaty.items():
        summaries[nr] = {
            "rok": year,
            "nr_automatu": nr,
            "obrot_brutto_ytd": round(sum(r.get("obrot_brutto", 0) for r in rows), 2),
            "prowizja_ytd": round(sum(r.get("prowizja_total", 0) for r in rows), 2),
            "koszty_ytd": round(sum(r.get("koszty_total", 0) for r in rows), 2),
            "wynik_netto_ytd": round(sum(r.get("wynik_netto", 0) for r in rows), 2),
            "liczba_miesiecy": len({r.get("miesiac") for r in rows}),
        }
    return summaries


# ---------------------------------------------------------------------------
# Main report assembly
# ---------------------------------------------------------------------------

def assemble_report_row(
    info: dict,
    obrot: dict | None,
    costs: dict | None,
    cfg: dict,
    info_changes: list[str],
    year: int,
    month: int,
    prowizja: dict | None = None,
) -> dict:
    """Build a single report row (all 8 columns) for one automat.

    The *prowizja* argument accepts a pre-loaded commission record (e.g. fetched
    from the database via load_prowizja).  When None the commission is calculated
    on-the-fly from config rates, preserving backwards compatibility.
    """
    nr = info["nr_automatu"]
    obrot_brutto = float(obrot.get("obrot_brutto", 0)) if obrot else 0.0
    liczba_transakcji = int(obrot.get("liczba_transakcji", 0)) if obrot else 0
    srednia_transakcja = float(obrot.get("srednia_transakcja", 0)) if obrot else 0.0

    # 3-PROWIZJA – use pre-loaded record or fall back to calculation
    if prowizja is not None:
        commission = prowizja
    else:
        commission = calculate_commission(
            obrot_brutto,
            info.get("przewoznik", ""),
            cfg.get("commission_rates", {}),
        )

    # 4-PODATEK
    tax = calculate_tax(obrot_brutto, cfg.get("tax", {}))

    # 5-KOSZTY
    cost_record = costs or {}
    koszty_total = aggregate_costs(cost_record)
    uwagi_koszty = cost_record.get("uwagi_koszty", "")

    # 6-PO SUMIE
    wynik = calculate_wynik(
        tax["obrot_netto"],
        obrot_brutto,
        commission["prowizja_total"],
        koszty_total,
        cfg.get("alerts", {}),
    )

    # 8-UWAGI – automatic entries
    auto_uwagi_parts = []
    if info_changes:
        auto_uwagi_parts.append("INFO ZMIANA: " + "; ".join(info_changes))
    if not obrot:
        auto_uwagi_parts.append("BRAK DANYCH OBROTU")
    if not costs:
        auto_uwagi_parts.append("BRAK DANYCH KOSZTÓW")
    if wynik.get("alert_wynik"):
        auto_uwagi_parts.append(wynik["alert_wynik"])
    if uwagi_koszty:
        auto_uwagi_parts.append(f"Koszty: {uwagi_koszty}")

    uwagi = " | ".join(auto_uwagi_parts)

    row: dict = {
        # Metadata
        "rok": year,
        "miesiac": month,
        # 1-INFO
        "nr_automatu": nr,
        "lokalizacja": info.get("lokalizacja", ""),
        "przewoznik": info.get("przewoznik", ""),
        "operator": info.get("operator", ""),
        "segment": info.get("segment", ""),
        "status": info.get("status", ""),
        "data_instalacji": info.get("data_instalacji", ""),
        # 2-OBRÓT
        "obrot_brutto": obrot_brutto,
        "liczba_transakcji": liczba_transakcji,
        "srednia_transakcja": srednia_transakcja,
        # 3-PROWIZJA
        "prowizja_elavon": commission["prowizja_elavon"],
        "prowizja_interchange": commission["prowizja_interchange"],
        "prowizja_total": commission["prowizja_total"],
        "stawka_elavon_pct": commission["stawka_elavon_pct"],
        "stawka_interchange_pct": commission["stawka_interchange_pct"],
        # 4-PODATEK
        "obrot_netto": tax["obrot_netto"],
        "vat_nalezny": tax["vat_nalezny"],
        "stawka_vat_pct": tax["stawka_vat_pct"],
        # 5-KOSZTY
        "koszt_czynsz": float(cost_record.get("koszt_czynsz", 0)),
        "koszt_prad": float(cost_record.get("koszt_prad", 0)),
        "koszt_elavon": float(cost_record.get("koszt_elavon", 0)),
        "koszt_poczta": float(cost_record.get("koszt_poczta", 0)),
        "koszt_amortyzacja": float(cost_record.get("koszt_amortyzacja", 0)),
        "koszt_serwis": float(cost_record.get("koszt_serwis", 0)),
        "koszt_transmisja": float(cost_record.get("koszt_transmisja", 0)),
        "koszt_pozostale": float(cost_record.get("koszt_pozostale", 0)),
        "koszty_total": koszty_total,
        # 6-PO SUMIE
        "wynik_netto": wynik["wynik_netto"],
        "wynik_brutto": wynik["wynik_brutto"],
        "marza_netto_pct": wynik["marza_netto_pct"],
        # 8-UWAGI
        "uwagi": uwagi,
    }
    return row


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

COLUMN_GROUPS = {
    "1-INFO": [
        "nr_automatu", "lokalizacja", "przewoznik", "operator",
        "segment", "status", "data_instalacji",
    ],
    "2-OBRÓT": ["obrot_brutto", "liczba_transakcji", "srednia_transakcja"],
    "3-PROWIZJA": [
        "prowizja_elavon", "prowizja_interchange", "prowizja_total",
        "stawka_elavon_pct", "stawka_interchange_pct",
    ],
    "4-PODATEK": ["obrot_netto", "vat_nalezny", "stawka_vat_pct"],
    "5-KOSZTY": [
        "koszt_czynsz", "koszt_prad", "koszt_elavon", "koszt_poczta",
        "koszt_amortyzacja", "koszt_serwis", "koszt_transmisja",
        "koszt_pozostale", "koszty_total",
    ],
    "6-PO SUMIE": ["wynik_netto", "wynik_brutto", "marza_netto_pct"],
    "7-PODSUMOWANIE ROCZNE": [
        "obrot_brutto_ytd", "prowizja_ytd", "koszty_ytd",
        "wynik_netto_ytd", "liczba_miesiecy",
    ],
    "8-UWAGI": ["uwagi"],
}

CURRENCY_COLS = {
    "obrot_brutto", "obrot_netto", "vat_nalezny",
    "prowizja_elavon", "prowizja_interchange", "prowizja_total",
    "koszt_czynsz", "koszt_prad", "koszt_elavon", "koszt_poczta",
    "koszt_amortyzacja", "koszt_serwis", "koszt_transmisja",
    "koszt_pozostale", "koszty_total",
    "wynik_netto", "wynik_brutto",
    "obrot_brutto_ytd", "prowizja_ytd", "koszty_ytd", "wynik_netto_ytd",
    "srednia_transakcja",
}

PERCENT_COLS = {"stawka_elavon_pct", "stawka_interchange_pct", "stawka_vat_pct", "marza_netto_pct"}


def export_to_excel(
    main_rows: list[dict],
    annual_summary: dict[str, dict],
    year: int,
    month: int,
    output_dir: Path,
    filename_template: str,
) -> Path:
    """Export the report to an Excel workbook with separate sheets per section."""
    if not _PANDAS_AVAILABLE:
        raise RuntimeError("Pakiet pandas i openpyxl są wymagane do eksportu Excel.")

    output_dir.mkdir(parents=True, exist_ok=True)
    filename = filename_template.format(year=year, month=month)
    out_path = output_dir / filename

    df_main = pd.DataFrame(main_rows)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Sheet: Raport Główny (all columns 1-6 + 8)
        main_cols = ["rok", "miesiac"]
        for group in ["1-INFO", "2-OBRÓT", "3-PROWIZJA", "4-PODATEK", "5-KOSZTY", "6-PO SUMIE", "8-UWAGI"]:
            main_cols.extend(COLUMN_GROUPS[group])
        available_cols = [c for c in main_cols if c in df_main.columns]
        df_main[available_cols].to_excel(writer, sheet_name="Raport Główny", index=False)
        _format_excel_sheet(writer, "Raport Główny", available_cols)

        # Sheet: 7-Roczne
        if annual_summary:
            df_annual = pd.DataFrame(list(annual_summary.values()))
            annual_cols = ["rok", "nr_automatu"] + COLUMN_GROUPS["7-PODSUMOWANIE ROCZNE"]
            avail_annual = [c for c in annual_cols if c in df_annual.columns]
            df_annual[avail_annual].to_excel(writer, sheet_name="7-Roczne", index=False)
            _format_excel_sheet(writer, "7-Roczne", avail_annual)

    log.info("Eksport Excel zakończony: %s", out_path)
    return out_path


def _format_excel_sheet(writer, sheet_name: str, columns: list[str]) -> None:
    """Apply basic formatting to an Excel sheet."""
    try:
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter

        ws = writer.sheets[sheet_name]
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)

        for col_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

            # Column width
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = max(len(col_name) + 4, 14)

            # Number format for data rows
            for row_idx in range(2, ws.max_row + 1):
                data_cell = ws.cell(row=row_idx, column=col_idx)
                if col_name in CURRENCY_COLS:
                    data_cell.number_format = '#,##0.00 "PLN"'
                elif col_name in PERCENT_COLS:
                    data_cell.number_format = '0.00"%"'

        ws.freeze_panes = "A2"
    except Exception as exc:
        log.warning("Nie udało się sformatować arkusza '%s': %s", sheet_name, exc)


# ---------------------------------------------------------------------------
# CSV audit export
# ---------------------------------------------------------------------------

def export_to_csv(
    main_rows: list[dict],
    year: int,
    month: int,
    output_dir: Path,
    filename_template: str,
) -> Path:
    """Export all rows to a CSV audit file."""
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = filename_template.format(year=year, month=month)
    out_path = output_dir / filename

    if not main_rows:
        log.warning("Brak danych do eksportu CSV.")
        out_path.write_text("", encoding="utf-8")
        return out_path

    fieldnames = list(main_rows[0].keys())
    with out_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(main_rows)

    log.info("Eksport CSV audytowy zakończony: %s", out_path)
    return out_path


# ---------------------------------------------------------------------------
# Email notification
# ---------------------------------------------------------------------------

def _build_kpi_summary(main_rows: list[dict], year: int, month: int) -> dict:
    """Build KPI summary for notification."""
    return {
        "okres": f"{year}-{month:02d}",
        "liczba_automatow": len(main_rows),
        "obrot_total": round(sum(r.get("obrot_brutto", 0) for r in main_rows), 2),
        "prowizja_total": round(sum(r.get("prowizja_total", 0) for r in main_rows), 2),
        "koszty_total": round(sum(r.get("koszty_total", 0) for r in main_rows), 2),
        "wynik_netto_total": round(sum(r.get("wynik_netto", 0) for r in main_rows), 2),
    }


def send_email_notification(
    cfg: dict,
    main_rows: list[dict],
    year: int,
    month: int,
    excel_path: Path,
    status: str = "OK",
) -> None:
    """Send email notification with KPI summary and Excel attachment."""
    email_cfg = cfg.get("email", {})
    if not email_cfg.get("enabled", False):
        log.info("Powiadomienia e-mail wyłączone w konfiguracji.")
        return

    kpi = _build_kpi_summary(main_rows, year, month)
    subject = email_cfg["subject"].format(year=year, month=month)

    body = (
        f"Raport TVM P&L za {kpi['okres']}\n"
        f"Status: {status}\n\n"
        f"KPI:\n"
        f"  Liczba automatów:  {kpi['liczba_automatow']}\n"
        f"  Obrót brutto:      {kpi['obrot_total']:,.2f} PLN\n"
        f"  Prowizja:          {kpi['prowizja_total']:,.2f} PLN\n"
        f"  Koszty:            {kpi['koszty_total']:,.2f} PLN\n"
        f"  Wynik netto:       {kpi['wynik_netto_total']:,.2f} PLN\n"
    )

    msg = MIMEMultipart()
    msg["From"] = email_cfg["from_address"]
    msg["To"] = ", ".join(email_cfg.get("to_addresses", []))
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    # Attachment
    if excel_path and excel_path.exists():
        from email.mime.base import MIMEBase
        from email import encoders as email_encoders

        with excel_path.open("rb") as fh:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(fh.read())
        email_encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={excel_path.name}",
        )
        msg.attach(part)

    password = (
        os.environ.get("EMAIL_PASSWORD")
        or email_cfg.get("from_password", "")
    )
    smtp_host = email_cfg["smtp_host"]
    smtp_port = int(email_cfg.get("smtp_port", 587))
    use_tls = email_cfg.get("smtp_use_tls", True)

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as server:
            if use_tls:
                server.starttls()
            if password:
                server.login(email_cfg["from_address"], password)
            server.sendmail(
                email_cfg["from_address"],
                email_cfg.get("to_addresses", []),
                msg.as_string(),
            )
        log.info("Powiadomienie e-mail wysłane do: %s", email_cfg.get("to_addresses"))
    except Exception as exc:
        log.error("Błąd wysyłania e-mail: %s", exc)


def send_teams_notification(
    cfg: dict,
    main_rows: list[dict],
    year: int,
    month: int,
    status: str = "OK",
) -> None:
    """Send Microsoft Teams webhook notification with KPI card."""
    teams_cfg = cfg.get("teams", {})
    if not teams_cfg.get("enabled", False):
        log.info("Powiadomienia Teams wyłączone w konfiguracji.")
        return
    if not _REQUESTS_AVAILABLE:
        log.error("Pakiet 'requests' nie jest zainstalowany – Teams notification pominięty.")
        return

    kpi = _build_kpi_summary(main_rows, year, month)
    webhook_url = teams_cfg.get("webhook_url", "")
    if not webhook_url:
        log.error("Brak webhook_url w konfiguracji Teams.")
        return

    card = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": "1F4E79",
        "summary": f"Raport TVM P&L {kpi['okres']}",
        "sections": [
            {
                "activityTitle": f"📊 Raport TVM P&L {kpi['okres']}",
                "activitySubtitle": f"Status: {status}",
                "facts": [
                    {"name": "Liczba automatów", "value": str(kpi["liczba_automatow"])},
                    {"name": "Obrót brutto", "value": f"{kpi['obrot_total']:,.2f} PLN"},
                    {"name": "Prowizja", "value": f"{kpi['prowizja_total']:,.2f} PLN"},
                    {"name": "Koszty", "value": f"{kpi['koszty_total']:,.2f} PLN"},
                    {"name": "Wynik netto", "value": f"{kpi['wynik_netto_total']:,.2f} PLN"},
                ],
            }
        ],
    }

    try:
        response = _requests.post(webhook_url, json=card, timeout=15)
        response.raise_for_status()
        log.info("Powiadomienie Teams wysłane.")
    except Exception as exc:
        log.error("Błąd wysyłania powiadomienia Teams: %s", exc)


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------

def run_report(cfg: dict, year: int, month: int) -> int:
    """Main orchestration: fetch data, calculate, export, notify.

    Returns exit code (0 = success, 1 = error).
    """
    log.info("=" * 60)
    log.info("Raport TVM P&L – okres: %d-%02d", year, month)
    log.info("=" * 60)

    output_cfg = cfg.get("output", {})
    output_dir = Path(output_cfg.get("directory", "output"))
    snapshots_dir = Path(output_cfg.get("snapshots_directory", "snapshots"))

    status = "OK"
    excel_path: Path | None = None

    try:
        engine = get_db_engine(cfg)

        # 1-INFO + change tracking
        info_records, info_changes = process_info(engine, year, month, snapshots_dir)
        info_map = {r["nr_automatu"]: r for r in info_records}

        # 2-OBRÓT
        table_trans = cfg["database"]["tables"]["transakcje"]
        obrot_records = fetch_obrot_from_db(engine, year, month, table_trans)
        obrot_map = {r["nr_automatu"]: r for r in obrot_records}

        # 5-KOSZTY
        costs_map = load_costs(cfg, engine, year, month)

        # 3-PROWIZJA – via adapter (calculated from rates or fetched from DB)
        prowizja_map = load_prowizja(cfg, engine, year, month, info_map, obrot_map)

        # Assemble rows
        main_rows: list[dict] = []
        for nr, info in info_map.items():
            row = assemble_report_row(
                info=info,
                obrot=obrot_map.get(nr),
                costs=costs_map.get(nr),
                cfg=cfg,
                info_changes=info_changes.get(nr, []),
                year=year,
                month=month,
                prowizja=prowizja_map.get(nr),
            )
            main_rows.append(row)

        log.info("Przetworzono %d automatów.", len(main_rows))

        # 7-PODSUMOWANIE ROCZNE
        annual_summary = build_annual_summary_per_automat(main_rows, year)

        # Export Excel
        excel_path = export_to_excel(
            main_rows=main_rows,
            annual_summary=annual_summary,
            year=year,
            month=month,
            output_dir=output_dir,
            filename_template=output_cfg.get(
                "excel_filename", "TVM_PL_{year}_{month:02d}.xlsx"
            ),
        )

        # Export CSV audit
        export_to_csv(
            main_rows=main_rows,
            year=year,
            month=month,
            output_dir=output_dir,
            filename_template=output_cfg.get(
                "csv_audit_filename", "TVM_PL_{year}_{month:02d}_audit.csv"
            ),
        )

    except Exception as exc:
        log.error("Błąd podczas generowania raportu: %s", exc, exc_info=True)
        status = f"BŁĄD: {exc}"
        main_rows = []

    # Notifications (always attempt, even on partial failure)
    send_email_notification(cfg, main_rows, year, month, excel_path or Path(), status)
    send_teams_notification(cfg, main_rows, year, month, status)

    log.info("Raport zakończony. Status: %s", status)
    return 0 if status == "OK" else 1


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Miesięczne zestawienie TVM P&L wg schematu 1-8.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Przykłady:\n"
            "  python export_automaty.py\n"
            "  python export_automaty.py --month 2026-02\n"
            "  python export_automaty.py --month 2026-02 --config config.local.yaml\n"
        ),
    )
    parser.add_argument(
        "--month",
        metavar="YYYY-MM",
        help="Okres raportowy (domyślnie: poprzedni pełny miesiąc).",
    )
    parser.add_argument(
        "--config",
        metavar="PATH",
        default=str(DEFAULT_CONFIG_PATH),
        help="Ścieżka do pliku konfiguracyjnego YAML.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    cfg = load_config(args.config)
    year, month = resolve_reporting_period(args.month, cfg)
    return run_report(cfg, year, month)


if __name__ == "__main__":
    sys.exit(main())
