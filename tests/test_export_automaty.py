"""
tests/test_export_automaty.py – testy jednostkowe dla modułu export_automaty.

Testują logikę kalkulacji bez potrzeby połączenia z bazą danych,
plikiem Excel ani zewnętrznymi serwisami.
"""

import json
import tempfile
import csv
from datetime import date
from pathlib import Path

import pytest

# Importujemy testowane funkcje
from export_automaty import (
    _parse_yyyymm,
    _validate_sql_identifier,
    aggregate_costs,
    assemble_report_row,
    build_annual_summary,
    build_annual_summary_per_automat,
    build_connection_string,
    calculate_commission,
    calculate_tax,
    calculate_wynik,
    detect_info_changes,
    load_costs_from_csv,
    load_snapshot,
    month_date_range,
    resolve_reporting_period,
    save_snapshot,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def make_cfg(**overrides) -> dict:
    """Minimal valid config dict."""
    cfg = {
        "reporting": {"default_period": "auto"},
        "database": {
            "type": "mssql",
            "host": "localhost",
            "port": 1433,
            "name": "tvm_db",
            "user": "user",
            "password": "pass",
            "tables": {
                "automaty": "automaty",
                "transakcje": "transakcje",
                "koszty": "tvm_koszty",
            },
        },
        "costs": {"source": "csv", "csv_path": "kas.csv"},
        "commission_rates": {
            "elavon_default": 0.0125,
            "interchange_per_carrier": {
                "PKP IC": 0.0080,
                "default": 0.0100,
            },
        },
        "tax": {"vat_rate": 0.23, "method": "from_gross"},
        "alerts": {"wynik_netto_min": -500.0},
        "output": {
            "directory": "output",
            "excel_filename": "TVM_PL_{year}_{month:02d}.xlsx",
            "csv_audit_filename": "TVM_PL_{year}_{month:02d}_audit.csv",
            "snapshots_directory": "snapshots",
        },
        "email": {"enabled": False},
        "teams": {"enabled": False},
    }
    cfg.update(overrides)
    return cfg


SAMPLE_INFO = {
    "nr_automatu": "TVM-001",
    "lokalizacja": "Gdańsk Główny",
    "przewoznik": "PKP IC",
    "operator": "ELAVON",
    "segment": "Nord",
    "status": "AKTYWNY",
    "data_instalacji": "2022-01-15",
}

SAMPLE_OBROT = {
    "nr_automatu": "TVM-001",
    "obrot_brutto": 10000.0,
    "liczba_transakcji": 200,
    "srednia_transakcja": 50.0,
}

SAMPLE_COSTS = {
    "koszt_czynsz": 1200.0,
    "koszt_prad": 85.5,
    "koszt_elavon": 49.0,
    "koszt_poczta": 12.0,
    "koszt_amortyzacja": 250.0,
    "koszt_serwis": 0.0,
    "koszt_transmisja": 35.0,
    "koszt_pozostale": 0.0,
    "uwagi_koszty": "",
}


# ---------------------------------------------------------------------------
# _parse_yyyymm
# ---------------------------------------------------------------------------

class TestParseYyyymm:
    def test_valid(self):
        assert _parse_yyyymm("2026-02") == (2026, 2)

    def test_valid_with_whitespace(self):
        assert _parse_yyyymm("  2025-12  ") == (2025, 12)

    def test_invalid_format_raises(self):
        with pytest.raises(ValueError, match="YYYY-MM"):
            _parse_yyyymm("2026/02")

    def test_invalid_month_raises(self):
        with pytest.raises(ValueError):
            _parse_yyyymm("2026-13")


# ---------------------------------------------------------------------------
# resolve_reporting_period
# ---------------------------------------------------------------------------

class TestResolveReportingPeriod:
    def test_cli_arg_takes_priority(self):
        cfg = make_cfg()
        assert resolve_reporting_period("2024-06", cfg) == (2024, 6)

    def test_config_explicit_period(self):
        cfg = make_cfg(reporting={"default_period": "2025-03"})
        assert resolve_reporting_period(None, cfg) == (2025, 3)

    def test_auto_returns_previous_month(self, monkeypatch):
        # Freeze today to 2026-03-09
        import export_automaty as mod
        monkeypatch.setattr(
            mod,
            "date",
            type("FakeDate", (), {"today": staticmethod(lambda: date(2026, 3, 9))})(),
        )
        assert resolve_reporting_period(None, make_cfg()) == (2026, 2)

    def test_auto_january_wraps_to_december(self, monkeypatch):
        import export_automaty as mod
        monkeypatch.setattr(
            mod,
            "date",
            type("FakeDate", (), {"today": staticmethod(lambda: date(2026, 1, 15))})(),
        )
        assert resolve_reporting_period(None, make_cfg()) == (2025, 12)


# ---------------------------------------------------------------------------
# month_date_range
# ---------------------------------------------------------------------------

class TestMonthDateRange:
    def test_february_non_leap(self):
        first, last = month_date_range(2026, 2)
        assert first == date(2026, 2, 1)
        assert last == date(2026, 2, 28)

    def test_february_leap(self):
        _, last = month_date_range(2024, 2)
        assert last == date(2024, 2, 29)

    def test_december(self):
        first, last = month_date_range(2026, 12)
        assert first == date(2026, 12, 1)
        assert last == date(2026, 12, 31)


# ---------------------------------------------------------------------------
# calculate_commission
# ---------------------------------------------------------------------------

class TestCalculateCommission:
    def test_pkpic_rates(self):
        cfg = make_cfg()["commission_rates"]
        result = calculate_commission(10000.0, "PKP IC", cfg)
        assert result["prowizja_elavon"] == 125.0       # 1.25%
        assert result["prowizja_interchange"] == 80.0   # 0.80%
        assert result["prowizja_total"] == 205.0

    def test_default_carrier_rate(self):
        cfg = make_cfg()["commission_rates"]
        result = calculate_commission(10000.0, "NIEZNANY", cfg)
        assert result["prowizja_interchange"] == 100.0  # 1.00% default
        assert result["prowizja_total"] == 225.0

    def test_zero_obrot(self):
        cfg = make_cfg()["commission_rates"]
        result = calculate_commission(0.0, "PKP IC", cfg)
        assert result["prowizja_total"] == 0.0

    def test_rounding(self):
        cfg = make_cfg()["commission_rates"]
        result = calculate_commission(333.33, "PKP IC", cfg)
        assert result["prowizja_elavon"] == round(333.33 * 0.0125, 2)
        assert result["prowizja_interchange"] == round(333.33 * 0.008, 2)


# ---------------------------------------------------------------------------
# calculate_tax
# ---------------------------------------------------------------------------

class TestCalculateTax:
    def test_from_gross_method(self):
        result = calculate_tax(12300.0, {"vat_rate": 0.23, "method": "from_gross"})
        # VAT = 12300 * 0.23 / 1.23 ≈ 2300.00
        assert result["vat_nalezny"] == round(12300.0 * 0.23 / 1.23, 2)
        assert result["obrot_netto"] == round(12300.0 - result["vat_nalezny"], 2)
        assert result["stawka_vat_pct"] == 23.0

    def test_to_net_method(self):
        result = calculate_tax(10000.0, {"vat_rate": 0.23, "method": "to_net"})
        assert result["vat_nalezny"] == 2300.0
        assert result["obrot_netto"] == 7700.0

    def test_zero_obrot(self):
        result = calculate_tax(0.0, {"vat_rate": 0.23, "method": "from_gross"})
        assert result["vat_nalezny"] == 0.0
        assert result["obrot_netto"] == 0.0


# ---------------------------------------------------------------------------
# aggregate_costs
# ---------------------------------------------------------------------------

class TestAggregateCosts:
    def test_sum_all_fields(self):
        total = aggregate_costs(SAMPLE_COSTS)
        expected = 1200 + 85.5 + 49 + 12 + 250 + 0 + 35 + 0
        assert total == round(expected, 2)

    def test_empty_record(self):
        assert aggregate_costs({}) == 0.0

    def test_missing_fields_treated_as_zero(self):
        assert aggregate_costs({"koszt_czynsz": 100.0}) == 100.0


# ---------------------------------------------------------------------------
# calculate_wynik
# ---------------------------------------------------------------------------

class TestCalculateWynik:
    def test_positive_result(self):
        result = calculate_wynik(
            obrot_netto=8130.08,
            obrot_brutto=10000.0,
            prowizja_total=205.0,
            koszty_total=1631.5,
            alerts_cfg={"wynik_netto_min": -500.0},
        )
        assert result["wynik_netto"] == round(8130.08 - 205.0 - 1631.5, 2)
        assert result["alert_wynik"] == ""

    def test_negative_result_triggers_alert(self):
        result = calculate_wynik(
            obrot_netto=100.0,
            obrot_brutto=130.0,
            prowizja_total=50.0,
            koszty_total=2000.0,
            alerts_cfg={"wynik_netto_min": -500.0},
        )
        assert result["wynik_netto"] < -500
        assert "ALERT" in result["alert_wynik"]

    def test_zero_obrot_netto_margin(self):
        result = calculate_wynik(0.0, 0.0, 0.0, 0.0, {})
        assert result["marza_netto_pct"] == 0.0


# ---------------------------------------------------------------------------
# detect_info_changes
# ---------------------------------------------------------------------------

class TestDetectInfoChanges:
    def test_no_changes(self):
        prev = {"TVM-001": dict(SAMPLE_INFO)}
        changes = detect_info_changes([SAMPLE_INFO], prev)
        assert "TVM-001" not in changes

    def test_location_change_detected(self):
        prev = {"TVM-001": {**SAMPLE_INFO, "lokalizacja": "Gdańsk Stocznia"}}
        changes = detect_info_changes([SAMPLE_INFO], prev)
        assert "TVM-001" in changes
        assert any("lokalizacja" in c for c in changes["TVM-001"])

    def test_new_automat_detected(self):
        changes = detect_info_changes([SAMPLE_INFO], {})
        assert "TVM-001" in changes
        assert changes["TVM-001"] == ["NOWY AUTOMAT"]

    def test_multiple_fields_changed(self):
        prev = {
            "TVM-001": {
                **SAMPLE_INFO,
                "lokalizacja": "Stara Lokalizacja",
                "status": "SERWIS",
            }
        }
        changes = detect_info_changes([SAMPLE_INFO], prev)
        assert len(changes["TVM-001"]) == 2


# ---------------------------------------------------------------------------
# save_snapshot / load_snapshot
# ---------------------------------------------------------------------------

class TestSnapshot:
    def test_save_and_load(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            snapshots_dir = Path(tmpdir)
            records = [SAMPLE_INFO]
            save_snapshot(snapshots_dir, 2026, 2, records)

            loaded = load_snapshot(snapshots_dir, 2026, 2)
            assert "TVM-001" in loaded
            assert loaded["TVM-001"]["lokalizacja"] == "Gdańsk Główny"

    def test_load_missing_snapshot_returns_empty(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            result = load_snapshot(Path(tmpdir), 2026, 1)
            assert result == {}


# ---------------------------------------------------------------------------
# load_costs_from_csv
# ---------------------------------------------------------------------------

class TestLoadCostsFromCsv:
    def _write_csv(self, path: Path, rows: list[dict]) -> None:
        with path.open("w", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
            writer.writeheader()
            writer.writerows(rows)

    def test_loads_correct_month(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / "kas.csv"
            self._write_csv(csv_path, [
                {
                    "nr_automatu": "TVM-001", "rok": 2026, "miesiac": 2,
                    "koszt_czynsz": 1200, "koszt_prad": 85.5,
                    "koszt_elavon": 49, "koszt_poczta": 12,
                    "koszt_amortyzacja": 250, "koszt_serwis": 0,
                    "koszt_transmisja": 35, "koszt_pozostale": 0,
                    "uwagi_koszty": "",
                }
            ])
            result = load_costs_from_csv(csv_path, 2026, 2)
            assert "TVM-001" in result
            assert result["TVM-001"]["koszt_czynsz"] == 1200.0

    def test_ignores_other_months(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / "kas.csv"
            self._write_csv(csv_path, [
                {
                    "nr_automatu": "TVM-001", "rok": 2026, "miesiac": 1,
                    "koszt_czynsz": 999, "koszt_prad": 0, "koszt_elavon": 0,
                    "koszt_poczta": 0, "koszt_amortyzacja": 0,
                    "koszt_serwis": 0, "koszt_transmisja": 0,
                    "koszt_pozostale": 0, "uwagi_koszty": "",
                }
            ])
            result = load_costs_from_csv(csv_path, 2026, 2)
            assert result == {}

    def test_missing_file_returns_empty(self):
        result = load_costs_from_csv(Path("/nonexistent/kas.csv"), 2026, 2)
        assert result == {}


# ---------------------------------------------------------------------------
# assemble_report_row
# ---------------------------------------------------------------------------

class TestAssembleReportRow:
    def test_all_sections_present(self):
        cfg = make_cfg()
        row = assemble_report_row(
            info=SAMPLE_INFO,
            obrot=SAMPLE_OBROT,
            costs=SAMPLE_COSTS,
            cfg=cfg,
            info_changes=[],
            year=2026,
            month=2,
        )
        # 1-INFO fields
        assert row["nr_automatu"] == "TVM-001"
        assert row["przewoznik"] == "PKP IC"
        # 2-OBRÓT
        assert row["obrot_brutto"] == 10000.0
        # 3-PROWIZJA
        assert row["prowizja_total"] == pytest.approx(205.0)
        # 4-PODATEK
        assert row["vat_nalezny"] == pytest.approx(10000 * 0.23 / 1.23, abs=0.01)
        # 5-KOSZTY
        assert row["koszty_total"] == pytest.approx(1631.5)
        # 6-PO SUMIE
        assert "wynik_netto" in row
        # 8-UWAGI (no warnings expected)
        assert row["uwagi"] == ""

    def test_missing_obrot_generates_warning(self):
        cfg = make_cfg()
        row = assemble_report_row(
            info=SAMPLE_INFO,
            obrot=None,
            costs=SAMPLE_COSTS,
            cfg=cfg,
            info_changes=[],
            year=2026,
            month=2,
        )
        assert row["obrot_brutto"] == 0.0
        assert "BRAK DANYCH OBROTU" in row["uwagi"]

    def test_missing_costs_generates_warning(self):
        cfg = make_cfg()
        row = assemble_report_row(
            info=SAMPLE_INFO,
            obrot=SAMPLE_OBROT,
            costs=None,
            cfg=cfg,
            info_changes=[],
            year=2026,
            month=2,
        )
        assert row["koszty_total"] == 0.0
        assert "BRAK DANYCH KOSZTÓW" in row["uwagi"]

    def test_info_change_in_uwagi(self):
        cfg = make_cfg()
        row = assemble_report_row(
            info=SAMPLE_INFO,
            obrot=SAMPLE_OBROT,
            costs=SAMPLE_COSTS,
            cfg=cfg,
            info_changes=["lokalizacja: Stara → Gdańsk Główny"],
            year=2026,
            month=2,
        )
        assert "INFO ZMIANA" in row["uwagi"]

    def test_low_result_alert_in_uwagi(self):
        cfg = make_cfg()
        # Very low obrot → wynik deeply negative
        low_obrot = {**SAMPLE_OBROT, "obrot_brutto": 0.0, "srednia_transakcja": 0.0}
        row = assemble_report_row(
            info=SAMPLE_INFO,
            obrot=low_obrot,
            costs=SAMPLE_COSTS,
            cfg=cfg,
            info_changes=[],
            year=2026,
            month=2,
        )
        assert "ALERT" in row["uwagi"]


# ---------------------------------------------------------------------------
# build_annual_summary
# ---------------------------------------------------------------------------

class TestBuildAnnualSummary:
    def _make_rows(self, year: int, months: list[int]) -> list[dict]:
        return [
            {
                "rok": year,
                "miesiac": m,
                "nr_automatu": "TVM-001",
                "obrot_brutto": 10000.0,
                "prowizja_total": 200.0,
                "koszty_total": 1500.0,
                "wynik_netto": 8300.0,
            }
            for m in months
        ]

    def test_ytd_sum(self):
        rows = self._make_rows(2026, [1, 2, 3])
        summary = build_annual_summary(rows, 2026)
        assert summary["obrot_brutto_ytd"] == 30000.0
        assert summary["wynik_netto_ytd"] == 24900.0
        assert summary["liczba_miesiecy"] == 3

    def test_different_year_excluded(self):
        rows = self._make_rows(2025, [11, 12]) + self._make_rows(2026, [1])
        summary = build_annual_summary(rows, 2026)
        assert summary["obrot_brutto_ytd"] == 10000.0
        assert summary["liczba_miesiecy"] == 1

    def test_per_automat_summary(self):
        rows = self._make_rows(2026, [1, 2])
        summaries = build_annual_summary_per_automat(rows, 2026)
        assert "TVM-001" in summaries
        assert summaries["TVM-001"]["obrot_brutto_ytd"] == 20000.0


# ---------------------------------------------------------------------------
# build_connection_string
# ---------------------------------------------------------------------------

class TestBuildConnectionString:
    def test_mssql(self):
        db_cfg = {
            "type": "mssql", "host": "srv", "port": 1433,
            "name": "db", "user": "u", "password": "p",
        }
        cs = build_connection_string(db_cfg)
        assert cs.startswith("mssql+pyodbc://")
        assert "srv:1433/db" in cs

    def test_mysql(self):
        db_cfg = {
            "type": "mysql", "host": "srv", "port": 3306,
            "name": "db", "user": "u", "password": "p",
        }
        cs = build_connection_string(db_cfg)
        assert cs.startswith("mysql+pymysql://")

    def test_postgresql(self):
        db_cfg = {
            "type": "postgresql", "host": "srv", "port": 5432,
            "name": "db", "user": "u", "password": "p",
        }
        cs = build_connection_string(db_cfg)
        assert cs.startswith("postgresql+psycopg2://")

    def test_sqlite(self):
        db_cfg = {
            "type": "sqlite", "name": "/tmp/test.db",
            "host": "", "port": "", "user": "", "password": "",
        }
        cs = build_connection_string(db_cfg)
        assert cs == "sqlite:////tmp/test.db"

    def test_unknown_db_type_raises(self):
        db_cfg = {
            "type": "oracle", "host": "", "port": "",
            "name": "", "user": "", "password": "",
        }
        with pytest.raises(ValueError, match="Nieznany typ bazy"):
            build_connection_string(db_cfg)

    def test_password_from_env(self, monkeypatch):
        monkeypatch.setenv("DB_PASSWORD", "secret123")
        db_cfg = {
            "type": "mysql", "host": "srv", "port": 3306,
            "name": "db", "user": "u", "password": "",
        }
        cs = build_connection_string(db_cfg)
        assert "secret123" in cs

    def test_empty_password_no_encoded_colon(self):
        db_cfg = {
            "type": "mysql", "host": "srv", "port": 3306,
            "name": "db", "user": "u", "password": "",
        }
        cs = build_connection_string(db_cfg)
        # Empty password should produce empty segment between user: and @
        assert "u:@" in cs


# ---------------------------------------------------------------------------
# _validate_sql_identifier
# ---------------------------------------------------------------------------

class TestValidateSqlIdentifier:
    def test_valid_identifier(self):
        assert _validate_sql_identifier("automaty") == "automaty"
        assert _validate_sql_identifier("tvm_koszty") == "tvm_koszty"

    def test_invalid_identifier_raises(self):
        with pytest.raises(ValueError, match="Nieprawidłowa nazwa"):
            _validate_sql_identifier("automaty; DROP TABLE users--")

    def test_identifier_with_space_raises(self):
        with pytest.raises(ValueError):
            _validate_sql_identifier("my table")
