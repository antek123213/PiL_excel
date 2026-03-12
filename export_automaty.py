# export_automaty.py - Miesięczne zestawienie TVM P&L (ETAP 1: kolumny 1-INFO, 2-OBRÓT)
import psycopg2
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
from dateutil.relativedelta import relativedelta
import json
import argparse
import time
from pathlib import Path

# === KONFIGURACJA ===
DB_CONFIG = {
    'host': '192.168.101.20',
    'port': 5432,
    'database': 'postgres',  # Zmienione - będzie automatycznie wykrywana właściwa baza
    'user': 'sprzedaz',
    'password': 'tVregNm5'
}

# Schemat AV zawiera wszystkie dane (moneystats, dictionary, aggregators, summaries)
SCHEMA_AV = 'AV'
SNAPSHOT_DIR = Path(__file__).parent / 'snapshots'
OUTPUT_DIR = Path(__file__).parent / 'output'


# === FUNKCJE POMOCNICZE: DIAGNOSTYKA BAZY DANYCH ===

def list_available_databases(host, port, user, password):
    """
    Listuje dostępne bazy danych na serwerze PostgreSQL.
    """
    try:
        conn = psycopg2.connect(
            host=host,
            port=port,
            database='postgres',  # Domyślna baza systemowa
            user=user,
            password=password,
            connect_timeout=10
        )
        cursor = conn.cursor()
        cursor.execute("SELECT datname FROM pg_database WHERE datistemplate = false ORDER BY datname;")
        databases = [row[0] for row in cursor.fetchall()]
        cursor.close()
        conn.close()
        return databases
    except psycopg2.Error as e:
        print(f"❌ Nie można pobrać listy baz danych: {e}")
        return []


def connect_to_db(database_name=None, max_retries=3, retry_delay=5):
    """
    Łączy się z bazą danych z mechanizmem retry.
    Jeśli database_name nie podane, automatycznie wykrywa bazę 'monitor'.
    """
    # Sprawdź dostępne bazy danych
    if database_name is None:
        print("\n🔍 Sprawdzanie dostępnych baz danych...")
        databases = list_available_databases(
            DB_CONFIG['host'],
            DB_CONFIG['port'],
            DB_CONFIG['user'],
            DB_CONFIG['password']
        )
        
        if databases:
            print(f"✓ Dostępne bazy danych: {', '.join(databases)}")
            
            # Szukaj bazy danych z nazwą "monitor"
            monitor_dbs = [db for db in databases if 'monitor' in db.lower()]
            if monitor_dbs:
                database_name = monitor_dbs[0]
                print(f"✓ Wybrano bazę danych: {database_name}")
            else:
                print(f"⚠ Brak bazy 'monitor'. Dostępne: {', '.join(databases)}")
                # Spróbuj pierwszej bazy z listy (pomijając postgres/template)
                non_system = [db for db in databases if db not in ['postgres', 'template0', 'template1']]
                if non_system:
                    database_name = non_system[0]
                    print(f"⚠ Używam pierwszej dostępnej bazy: {database_name}")
                else:
                    print("❌ Brak dostępnych baz danych użytkownika")
                    return None
        else:
            print("❌ Nie można pobrać listy baz danych")
            return None
    
    # Połącz się z wybraną bazą
    for attempt in range(1, max_retries + 1):
        try:
            conn = psycopg2.connect(
                host=DB_CONFIG['host'],
                port=DB_CONFIG['port'],
                database=database_name,
                user=DB_CONFIG['user'],
                password=DB_CONFIG['password'],
                connect_timeout=10
            )
            if attempt > 1:
                print(f"✓ Połączono z bazą '{database_name}' (próba {attempt}/{max_retries})")
            return conn
        except psycopg2.OperationalError as e:
            if attempt < max_retries:
                print(f"⚠ Próba {attempt}/{max_retries} nie powiodła się: {e}")
                print(f"  Ponowna próba za {retry_delay} sekund...")
                time.sleep(retry_delay)
            else:
                print(f"❌ Nie udało się połączyć z bazą '{database_name}' po {max_retries} próbach")
                raise
    return None


# === 1-INFO: SŁOWNIK AUTOMATÓW (AV.dictionary) ===

def get_dictionary_snapshot(conn):
    """
    Pobiera słownik automatów z AV.dictionary.
    Zwraca: dict {device_id: {value, description, groupid}}
    Filtruje tylko rzeczywiste ID automatów (numeric) lub w zakresie 1101-1299.
    """
    cursor = conn.cursor()
    query = """
    SELECT id, value, description, groupid
    FROM "AV".dictionary
    WHERE id::text ~ '^[0-9]+$'
    ORDER BY id
    """
    cursor.execute(query)
    rows = cursor.fetchall()
    cursor.close()
    
    snapshot = {}
    for row in rows:
        device_id, value, description, groupid = row
        # Konwertuj id do integer dla konsekwencji
        device_id_int = int(device_id)
        snapshot[device_id_int] = {
            'value': value,
            'description': description,
            'groupid': groupid
        }
    
    return snapshot


def load_previous_snapshot(month_str):
    """
    Ładuje poprzedni snapshot dictionary (miesiąc wcześniej).
    month_str: 'YYYY-MM'
    Zwraca: dict lub None
    Konwertuje klucze stringowe z JSON z powrotem na integery.
    """
    # Oblicz poprzedni miesiąc
    year, month = map(int, month_str.split('-'))
    prev_date = datetime(year, month, 1) - relativedelta(months=1)
    prev_month_str = prev_date.strftime('%Y-%m')
    
    snapshot_file = SNAPSHOT_DIR / f'dictionary_{prev_month_str}.json'
    if snapshot_file.exists():
        with open(snapshot_file, 'r', encoding='utf-8') as f:
            loaded = json.load(f)
            # Konwertuj string keys na integer keys
            return {int(k): v for k, v in loaded.items()}
    return None


def save_snapshot(snapshot, month_str):
    """
    Zapisuje snapshot dictionary do pliku JSON.
    """
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    snapshot_file = SNAPSHOT_DIR / f'dictionary_{month_str}.json'
    
    with open(snapshot_file, 'w', encoding='utf-8') as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2)
    
    print(f"✓ Snapshot zapisany: {snapshot_file}")


def compare_snapshots(current, previous):
    """
    Porównuje bieżący snapshot z poprzednim.
    Zwraca: dict {device_id: {status: 'new'|'changed'|'unchanged', data: {...}}}
    """
    result = {}
    
    for device_id, data in current.items():
        if device_id not in previous:
            result[device_id] = {'status': 'new', 'data': data}
        elif data != previous[device_id]:
            result[device_id] = {'status': 'changed', 'data': data}
        else:
            result[device_id] = {'status': 'unchanged', 'data': data}
    
    # Urządzenia usunięte (były w poprzednim, nie ma w bieżącym)
    for device_id in previous:
        if device_id not in current:
            result[device_id] = {'status': 'deleted', 'data': previous[device_id]}
    
    return result


# === 2-OBRÓT: DANE Z MONEYSTATS ZA MIESIĄC ===

def get_monthly_revenue(conn, month_str):
    """
    Pobiera obrót brutto per automat za dany miesiąc z tabeli AV.moneystats.
    month_str: 'YYYY-MM'
    Zwraca: dict {device_id: {'obrot_brutto_zl': float, 'liczba_transakcji': int}}
    """
    year, month = map(int, month_str.split('-'))
    start_date = datetime(year, month, 1)
    # Koniec miesiąca = początek następnego miesiąca
    end_date = start_date + relativedelta(months=1)
    
    start_str = start_date.strftime('%Y-%m-%d %H:%M:%S')
    end_str = end_date.strftime('%Y-%m-%d %H:%M:%S')
    
    cursor = conn.cursor()
    
    # Pobierz dane z AV.moneystats
    query = f"""
        SELECT 
            ms_deviceid AS device_id,
            SUM(COALESCE(ms_vtotal, 0)) AS obrot_brutto_zl,
            COUNT(*) AS liczba_transakcji
        FROM "{SCHEMA_AV}".moneystats
        WHERE ms_creadate >= '{start_str}'
          AND ms_creadate < '{end_str}'
          AND ms_deviceid BETWEEN 1101 AND 1299
        GROUP BY ms_deviceid
        ORDER BY device_id
    """
    
    cursor.execute(query)
    rows = cursor.fetchall()
    cursor.close()
    
    # Konwertuj na dict
    revenue_data = {}
    for row in rows:
        device_id, obrot, liczba = row
        revenue_data[device_id] = {
            'obrot_brutto_zl': float(obrot),
            'liczba_transakcji': int(liczba)
        }
    
    return revenue_data


# === EKSPORT DO EXCEL (KOLUMNY 1-2, RESZTA PLACEHOLDER) ===

def export_to_excel_PL(dictionary_comparison, revenue_data, month_str, filename):
    """
    Eksportuje raport P&L do Excela.
    Kolumny:
    1-INFO: nr_automatu, value, description, groupid, status_zmiany
    2-OBRÓT: obrot_brutto_zl, liczba_transakcji, przewoznik
    3-8: placeholder (do ETAPU 2 i 3)
    """
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    filepath = OUTPUT_DIR / filename
    
    wb = Workbook()
    ws = wb.active
    ws.title = f"P&L {month_str}"
    
    # === NAGŁÓWKI ===
    headers = [
        # 1-INFO
        'Nr automatu', 'Value', 'Description', 'GroupID', 'Status', 'Przewoźnik',
        # 2-OBRÓT
        'Obrót brutto (zł)', 'Liczba transakcji',
        # 3-PROWIZJA (placeholder)
        'Prowizja (zł)',
        # 4-RODZAJ TRANZAKCJI (placeholder)
        'Rodzaj transakcji',
        # 5-KOSZTY (placeholder)
        'Koszty (zł)',
        # 6-PO SUMIE (placeholder)
        'Wynik netto (zł)',
        # 7-PODSUMOWANIE ROCZNE (placeholder)
        'Suma roczna',
        # 8-UWAGI (placeholder)
        'Uwagi'
    ]
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # === DANE ===
    row_num = 2
    # Upewnij się, że wszystkie klucze są integerami
    all_device_ids = sorted(set(
        [int(k) for k in dictionary_comparison.keys()] + 
        [int(k) for k in revenue_data.keys()]
    ))
    
    for device_id in all_device_ids:
        # 1-INFO
        if device_id in dictionary_comparison:
            info = dictionary_comparison[device_id]
            status = info['status']
            data = info['data']
            value = data.get('value', '')
            description = data.get('description', '')
            groupid = data.get('groupid', '')
        else:
            value = description = groupid = ''
            status = 'brak w dict'
        
        # 2-OBRÓT
        if device_id in revenue_data:
            rev = revenue_data[device_id]
            obrot = rev['obrot_brutto_zl']
            liczba = rev['liczba_transakcji']
            przewoznik = ''  # TODO: Pobrać z tabeli aggregators
        else:
            obrot = 0.0
            liczba = 0
            przewoznik = ''
        
        # Placeholder dla kolumn 3-8
        prowizja = ''
        rodzaj = ''
        koszty = ''
        wynik_netto = ''
        suma_roczna = ''
        uwagi = ''
        
        # Wpisz dane do wiersza
        ws.cell(row=row_num, column=1).value = device_id
        ws.cell(row=row_num, column=2).value = value
        ws.cell(row=row_num, column=3).value = description
        ws.cell(row=row_num, column=4).value = groupid
        ws.cell(row=row_num, column=5).value = status
        ws.cell(row=row_num, column=6).value = przewoznik
        ws.cell(row=row_num, column=7).value = obrot
        ws.cell(row=row_num, column=7).number_format = '#,##0.00'
        ws.cell(row=row_num, column=8).value = liczba
        ws.cell(row=row_num, column=9).value = prowizja
        ws.cell(row=row_num, column=10).value = rodzaj
        ws.cell(row=row_num, column=11).value = koszty
        ws.cell(row=row_num, column=12).value = wynik_netto
        ws.cell(row=row_num, column=13).value = suma_roczna
        ws.cell(row=row_num, column=14).value = uwagi
        
        row_num += 1
    
    # Dostosuj szerokość kolumn
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column].width = adjusted_width
    
    wb.save(filepath)
    print(f"\n✓ Raport P&L eksportowany: {filepath}")
    print(f"✓ Liczba wierszy: {row_num - 2}")


# === MAIN ===

def main():
    parser = argparse.ArgumentParser(description='Miesięczne zestawienie TVM P&L')
    parser.add_argument(
        '--miesiac',
        type=str,
        help='Miesiąc raportu w formacie YYYY-MM (domyślnie: poprzedni miesiąc)',
        default=None
    )
    args = parser.parse_args()
    
    # Określ miesiąc raportowy
    if args.miesiac:
        month_str = args.miesiac
        # Walidacja formatu
        try:
            datetime.strptime(month_str, '%Y-%m')
        except ValueError:
            print(f"❌ Nieprawidłowy format miesiąca: {month_str}. Użyj formatu YYYY-MM")
            return
    else:
        # Domyślnie: poprzedni miesiąc
        today = datetime.now()
        prev_month = today - relativedelta(months=1)
        month_str = prev_month.strftime('%Y-%m')
    
    print(f"\n{'='*60}")
    print(f"  Miesięczne zestawienie TVM P&L - {month_str}")
    print(f"{'='*60}\n")
    
    # Połączenie z bazą (auto-detect bazy danych)
    try:
        conn = connect_to_db(database_name=None, max_retries=3, retry_delay=5)
        if conn is None:
            print("❌ Nie można nawiązać połączenia z bazą danych")
            return
    except psycopg2.Error as e:
        print(f"❌ Błąd połączenia z bazą: {e}")
        return
    
    # === KROK 1: Słownik automatów (AV.dictionary) ===
    print("\n[1/4] Pobieranie słownika automatów (AV.dictionary)...")
    current_snapshot = get_dictionary_snapshot(conn)
    print(f"✓ Pobrano {len(current_snapshot)} automatów ze słownika")
    
    # Porównaj z poprzednim miesiącem
    print("\n[2/4] Porównanie z poprzednim miesiącem...")
    previous_snapshot = load_previous_snapshot(month_str)
    
    if previous_snapshot is None:
        print("ℹ  Brak poprzedniego snapshot — pierwszy raport dla tego miesiąca")
        dictionary_comparison = {
            device_id: {'status': 'new', 'data': data}
            for device_id, data in current_snapshot.items()
        }
    else:
        dictionary_comparison = compare_snapshots(current_snapshot, previous_snapshot)
        new_count = sum(1 for v in dictionary_comparison.values() if v['status'] == 'new')
        changed_count = sum(1 for v in dictionary_comparison.values() if v['status'] == 'changed')
        unchanged_count = sum(1 for v in dictionary_comparison.values() if v['status'] == 'unchanged')
        deleted_count = sum(1 for v in dictionary_comparison.values() if v['status'] == 'deleted')
        
        print(f"✓ Nowe: {new_count}, Zmienione: {changed_count}, "
              f"Bez zmian: {unchanged_count}, Usunięte: {deleted_count}")
    
    # Zapisz snapshot dla bieżącego miesiąca
    save_snapshot(current_snapshot, month_str)
    
    # === KROK 2: Obrót miesięczny (moneystats) ===
    print(f"\n[3/4] Pobieranie obrotu za {month_str}...")
    revenue_data = get_monthly_revenue(conn, month_str)
    print(f"✓ Pobrano dane obrotu dla {len(revenue_data)} automatów")
    
    total_revenue = sum(v['obrot_brutto_zl'] for v in revenue_data.values())
    total_transactions = sum(v['liczba_transakcji'] for v in revenue_data.values())
    print(f"  Obrót brutto: {total_revenue:,.2f} zł")
    print(f"  Liczba transakcji: {total_transactions:,}")
    
    conn.close()
    
    # === KROK 3: Eksport do Excel ===
    print(f"\n[4/4] Eksport do Excel...")
    filename = f"PL_TVM_{month_str}.xlsx"
    export_to_excel_PL(dictionary_comparison, revenue_data, month_str, filename)
    
    print(f"\n{'='*60}")
    print(f"  ✓ RAPORT ZAKOŃCZONY POMYŚLNIE")
    print(f"  Kolumny zaimplementowane: 1-INFO, 2-OBRÓT")
    print(f"  Do zrobienia (ETAP 2): kolumny 3-6")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()