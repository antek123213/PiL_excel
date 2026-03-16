"""
Test sprawdza czy export_to_excel_PL tworzy poprawne nagłówki:
  Kolumna 1-INFO:    Nr automatu, Value, Description, GroupID, Status, Przewoźnik
  Kolumna 2-OBRÓT:   Obrót brutto (zł), Liczba transakcji
  Kolumna 3-PROWIZJA: Prowizja (zł)
    Kolumny 4 i 6 (Rodzaj transakcji, Wynik netto) są usunięte.
"""
import sys
import os
import pytest
from openpyxl import load_workbook

# Dodaj folder nadrzędny do ścieżki
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from export_automaty import export_to_excel_PL


MONTH = "2026-02"
FILENAME = f"TEST_PL_{MONTH}.xlsx"

SAMPLE_DICT = {
    1101: {"status": "unchanged", "data": {"value": "TVM1", "description": "Automat 1101", "groupid": "G1"}},
    1102: {"status": "new",       "data": {"value": "TVM2", "description": "Automat 1102", "groupid": "G1"}},
}

SAMPLE_REVENUE = {
    1101: {"obrot_brutto_zl": 1234.56, "liczba_transakcji": 42},
}


@pytest.fixture(scope="module")
def excel_path(tmp_path_factory):
    """Tworzy plik Excel w katalogu tymczasowym."""
    tmp = tmp_path_factory.mktemp("output")
    
    # Patch OUTPUT_DIR w module
    import export_automaty
    original_dir = export_automaty.OUTPUT_DIR
    export_automaty.OUTPUT_DIR = tmp

    export_to_excel_PL(SAMPLE_DICT, SAMPLE_REVENUE, MONTH, FILENAME)

    export_automaty.OUTPUT_DIR = original_dir
    return tmp / FILENAME


@pytest.fixture(scope="module")
def headers(excel_path):
    wb = load_workbook(excel_path)
    ws = wb.active
    return [ws.cell(row=1, column=i).value for i in range(1, ws.max_column + 1)]


# ── 1-INFO ─────────────────────────────────────────────────────────────────────
def test_col1_nr_automatu(headers):
    assert headers[0] == "Nr automatu", f"Kolumna 1: oczekiwano 'Nr automatu', got '{headers[0]}'"

def test_col2_value(headers):
    assert headers[1] == "Value", f"Kolumna 2: oczekiwano 'Value', got '{headers[1]}'"

def test_col3_description(headers):
    assert headers[2] == "Description", f"Kolumna 3: oczekiwano 'Description', got '{headers[2]}'"

def test_col4_groupid(headers):
    assert headers[3] == "GroupID", f"Kolumna 4: oczekiwano 'GroupID', got '{headers[3]}'"

def test_col5_status(headers):
    assert headers[4] == "Status", f"Kolumna 5: oczekiwano 'Status', got '{headers[4]}'"

def test_col6_przewoznik(headers):
    assert headers[5] == "Przewoźnik", f"Kolumna 6: oczekiwano 'Przewoźnik', got '{headers[5]}'"


# ── 2-OBRÓT ────────────────────────────────────────────────────────────────────
def test_col7_obrot(headers):
    assert headers[6] == "Obrót brutto (zł)", f"Kolumna 7: oczekiwano 'Obrót brutto (zł)', got '{headers[6]}'"

def test_col8_liczba_transakcji(headers):
    assert headers[7] == "Liczba transakcji", f"Kolumna 8: oczekiwano 'Liczba transakcji', got '{headers[7]}'"


# ── 3-PROWIZJA ─────────────────────────────────────────────────────────────────
def test_col9_prowizja(headers):
    assert headers[8] == "Prowizja (zł)", f"Kolumna 9: oczekiwano 'Prowizja (zł)', got '{headers[8]}'"


# ── 5,7,8 (po usunięciu 4 i 6) ───────────────────────────────────────────────
def test_col10_koszty(headers):
    assert headers[9] == "Koszty (zł)", f"Kolumna 10: oczekiwano 'Koszty (zł)', got '{headers[9]}'"

def test_col11_suma_roczna(headers):
    assert headers[10] == "Suma roczna", f"Kolumna 11: oczekiwano 'Suma roczna', got '{headers[10]}'"

def test_col12_uwagi(headers):
    assert headers[11] == "Uwagi", f"Kolumna 12: oczekiwano 'Uwagi', got '{headers[11]}'"

def test_removed_columns_are_not_present(headers):
    assert "Rodzaj transakcji" not in headers
    assert "Wynik netto (zł)" not in headers

def test_headers_count_after_column_removal(headers):
    assert len(headers) == 12


# ── Dane w wierszach ────────────────────────────────────────────────────────────
def test_data_device_id(excel_path):
    wb = load_workbook(excel_path)
    ws = wb.active
    device_ids = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    assert 1101 in device_ids
    assert 1102 in device_ids

def test_data_obrot(excel_path):
    wb = load_workbook(excel_path)
    ws = wb.active
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == 1101:
            assert ws.cell(row=row, column=7).value == pytest.approx(1234.56)
            break
    else:
        pytest.fail("Nie znaleziono automatu 1101 w danych")
