# export_automaty.py - Miesięczne zestawienie TVM P&L (ETAP 1: kolumny 1-INFO, 2-OBRÓT, 3-PROWIZJA)
import psycopg2
from psycopg2 import sql
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
from dateutil.relativedelta import relativedelta
import json
import argparse
import time
from pathlib import Path
import re
import calendar

# === KONFIGURACJA ===
DB_CONFIG = {
    'host': '192.168.101.20',
    'port': 5432,
    'database': 'postgres',  # Zmienione - będzie automatycznie wykrywana właściwa baza
    'user': 'sprzedaz',
    'password': 'tVregNm5',
    'sslmode': 'prefer'
}

# Domyślny schemat źródłowy danych operacyjnych.
SCHEMA = 'ARP'
SNAPSHOT_DIR = Path(__file__).parent / 'snapshots'
OUTPUT_DIR = Path(__file__).parent / 'output'
DEVICE_ID_RANGES = (
    (1101, 1141),
    (1201, 1299),
)

CARRIERS = ['IC', 'PR', 'KD', 'ARP', 'KW', 'KS', 'LKA', 'SKM', 'KMŁ']
TRANSACTIONS_SOURCE_CONFIG = {
    'table': 'transactions',
    'amount_col': 'fin_ptu_kwota',
    'date_col': 'fin_data_sp',
    'device_col': None,
}
TRANSACTIONS_DEVICE_COL_CANDIDATES = [
    'fin_nr_urz',
    'nr_urz',
    'tvm_id',
    'fin_tvm_id',
    'device_id',
    'id_automatu',
    'automat_id',
    'ms_deviceid',
]
TRANSACTIONS_CARRIER_COL_CANDIDATES = [
    'provider',
    'przewoznik',
    'carrier',
    'operator',
    'source',
    'payment_operator',
    'fin_przewoznik',
]

# Konfiguracja źródła prowizji - można nadpisać parametrami CLI
COMMISSION_SOURCE_CONFIG = {
    'schema': 'public',
    'table': 'provision',
    'device_col': 'tvm_id_prov',
    'amount_col': 'provision_prov',
    'date_col': 'valid_from',
    'date_to_col': 'valid_to',
    'provider_col': None,
    'provider_values': None
}


def _end_of_month_date(month_str):
    """
    Zwraca ostatni dzień miesiąca jako datetime (00:00:00).
    """
    year, month = map(int, month_str.split('-'))
    last_day = calendar.monthrange(year, month)[1]
    return datetime(year, month, last_day)


def build_output_filename(month_str, naming_mode='default'):
    """
    Buduje nazwę pliku wyjściowego.
    - default: PL_TVM_YYYY-MM.xlsx
    - monitor-style: PROVISIONYYYYMMDDHHMMSS.xlsx
    """
    if naming_mode == 'monitor-style':
        ts = datetime.now().strftime('%Y%m%d%H%M%S')
        return f"PROVISION{ts}.xlsx"
    return f"PL_TVM_{month_str}.xlsx"


def _build_device_filter_sql(column_sql):
    """
    Buduje warunek SQL filtrujący urządzenia wg ustalonych zakresów ID.
    """
    clauses = []
    for start_id, end_id in DEVICE_ID_RANGES:
        clauses.append(
            sql.SQL("({column} BETWEEN {start_id} AND {end_id})").format(
                column=column_sql,
                start_id=sql.Literal(start_id),
                end_id=sql.Literal(end_id),
            )
        )
    return sql.SQL("(") + sql.SQL(" OR ").join(clauses) + sql.SQL(")")


def extract_city_name(description):
    """
    Wyciąga nazwę miasta z pola description, usuwając numery automatów.
    Przykłady:
      'WROC-GL_1204'  -> 'WROC-GL'
      '1214_RZEPIN'   -> 'RZEPIN'
      'KRAK_GN_1263'  -> 'KRAK-GN'
      '2100779032'    -> ''  (czysto numeryczne)
      '1203: Opis'    -> ''  (placeholder)
    """
    if not description:
        return ''
    normalized = str(description).strip()
    if 'opis' in normalized.lower():
        return ''

    parts = normalized.split('_')
    letter_parts = [p for p in parts if re.search(r'[A-Za-z]', p)]
    result = '-'.join(letter_parts)
    if result in ('Opis', 'opis', ''):
        return ''
    return result


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
            sslmode=DB_CONFIG.get('sslmode', 'prefer'),
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


def database_has_required_tables(database_name, required_tables=None):
    """
    Sprawdza czy baza zawiera wymagane tabele raportowe.
    """
    if required_tables is None:
        required_tables = [
            (SCHEMA, 'dictionary'),
            (SCHEMA, 'moneystats'),
        ]

    try:
        conn = psycopg2.connect(
            host=DB_CONFIG['host'],
            port=DB_CONFIG['port'],
            database=database_name,
            user=DB_CONFIG['user'],
            password=DB_CONFIG['password'],
            sslmode=DB_CONFIG.get('sslmode', 'prefer'),
            connect_timeout=5
        )
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT table_schema, table_name
            FROM information_schema.tables
            WHERE table_type = 'BASE TABLE'
              AND table_schema = %s
              AND table_name = ANY(%s)
            """,
            (SCHEMA, [tbl for _, tbl in required_tables]),
        )
        existing = {(row[0], row[1]) for row in cursor.fetchall()}
        cursor.close()
        conn.close()
        return all(t in existing for t in required_tables)
    except psycopg2.Error:
        return False


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
                preferred_monitor = None
                for candidate_db in monitor_dbs:
                    if database_has_required_tables(candidate_db):
                        preferred_monitor = candidate_db
                        break

                if preferred_monitor:
                    database_name = preferred_monitor
                    print(f"✓ Wybrano bazę danych: {database_name} (z wymaganymi tabelami {SCHEMA})")
                else:
                    print(f"⚠ Baza monitor nie zawiera kompletu {SCHEMA}.dictionary + {SCHEMA}.moneystats")
                    print("  Szukanie alternatywnej bazy z wymaganymi tabelami...")
                    preferred_any = None
                    for candidate_db in databases:
                        if candidate_db in ['postgres', 'template0', 'template1']:
                            continue
                        if database_has_required_tables(candidate_db):
                            preferred_any = candidate_db
                            break

                    if preferred_any:
                        database_name = preferred_any
                        print(f"✓ Wybrano alternatywną bazę danych: {database_name}")
                    else:
                        database_name = monitor_dbs[0]
                        print(f"⚠ Używam bazy monitor bez pełnego zestawu tabel: {database_name}")
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
                sslmode=DB_CONFIG.get('sslmode', 'prefer'),
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


# === 1-INFO: SŁOWNIK AUTOMATÓW (ARP.dictionary) ===

def get_dictionary_snapshot(conn, schema=SCHEMA):
    """
    Pobiera słownik automatów z <schema>.dictionary.
    Zwraca: dict {device_id: {'description': description}}
    Filtruje tylko rzeczywiste ID automatów (numeric) z zakresów:
    1101-1141 oraz 1201-1299.
    """
    cursor = conn.cursor()
    query = sql.SQL(
        """
                SELECT id, description
                FROM {schema}.dictionary
                WHERE id::text ~ '^[0-9]+$'
                    AND {device_filter}
                ORDER BY id
                """
        ).format(
            schema=sql.Identifier(schema),
            device_filter=_build_device_filter_sql(sql.SQL("id::bigint")),
        )
    cursor.execute(query)
    rows = cursor.fetchall()
    cursor.close()
    
    snapshot = {}
    for row in rows:
        device_id, description = row
        snapshot[int(device_id)] = {
            'description': description,
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

def get_month_range(month_str):
    """
    Zwraca zakres dat dla miesiąca raportowego.
    month_str: 'YYYY-MM'
    """
    year, month = map(int, month_str.split('-'))
    start_date = datetime(year, month, 1)
    end_date = start_date + relativedelta(months=1)
    return start_date, end_date


def _pick_first_matching_column(columns, candidates):
    """
    Zwraca pierwszą kolumnę pasującą do listy kandydatów.
    Porównanie jest case-insensitive.
    """
    lowered = {c.lower(): c for c in columns}
    for candidate in candidates:
        if candidate.lower() in lowered:
            return lowered[candidate.lower()]
    return None


def discover_commission_source(conn, schema=SCHEMA):
    """
    Próbuje automatycznie znaleźć tabelę i kolumny prowizji w schema ARP.
    Zwraca dict ze źródłem lub None.
    """
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT table_name, column_name
        FROM information_schema.columns
        WHERE table_schema = %s
        ORDER BY table_name, ordinal_position
        """,
        (schema,)
    )
    rows = cursor.fetchall()
    cursor.close()

    if not rows:
        return None

    tables = {}
    for table_name, column_name in rows:
        tables.setdefault(table_name, []).append(column_name)

    device_candidates = [
        'device_id', 'deviceid', 'ms_deviceid', 'id_automatu', 'nr_automatu', 'automat_id'
    ]
    amount_candidates = [
        'commission', 'prowizja', 'prowizja_zl', 'fee', 'interchange',
        'elavon_fee', 'amount', 'kwota', 'kwota_prowizji'
    ]
    date_candidates = [
        'create_date', 'createdate', 'ms_creadate', 'data', 'date',
        'transaction_date', 'settlement_date'
    ]
    provider_candidates = [
        'provider', 'operator', 'aggregator', 'source', 'payment_operator'
    ]
    table_name_hints = ['prowiz', 'commission', 'interchange', 'elavon', 'summary', 'settlement']

    best_source = None
    best_score = -1

    for table_name, columns in tables.items():
        device_col = _pick_first_matching_column(columns, device_candidates)
        amount_col = _pick_first_matching_column(columns, amount_candidates)
        date_col = _pick_first_matching_column(columns, date_candidates)
        provider_col = _pick_first_matching_column(columns, provider_candidates)

        if not all([device_col, amount_col, date_col]):
            continue

        score = 0
        lowered_table = table_name.lower()
        score += sum(2 for token in table_name_hints if token in lowered_table)
        score += 3  # komplet kluczowych kolumn
        if provider_col:
            score += 1

        if score > best_score:
            best_score = score
            best_source = {
                'schema': schema,
                'table': table_name,
                'device_col': device_col,
                'amount_col': amount_col,
                'date_col': date_col,
                'provider_col': provider_col,
                'provider_values': None
            }

    return best_source


def get_monthly_commission(conn, month_str, source_config=None, strict_month_window=True):
    """
    Pobiera prowizję per automat za dany miesiąc.
    Zwraca: dict {device_id: {'prowizja_zl': float, 'liczba_rekordow': int}}
    """
    start_date, end_date = get_month_range(month_str)

    if source_config is None:
        source_config = COMMISSION_SOURCE_CONFIG

    source = dict(source_config)
    has_explicit_mapping = all([
        source.get('table'),
        source.get('device_col'),
        source.get('amount_col'),
        source.get('date_col')
    ])

    if not has_explicit_mapping:
        discovered = discover_commission_source(conn, schema=source.get('schema', SCHEMA))
        if discovered is None:
            print("⚠ Nie znaleziono źródła prowizji - kolumna 3-PROWIZJA pozostanie pusta")
            return {}
        source.update(discovered)
        print(
            "✓ Auto-detekcja prowizji: "
            f"{source['schema']}.{source['table']} "
            f"(device={source['device_col']}, amount={source['amount_col']}, date={source['date_col']})"
        )

    provider_col = source.get('provider_col')
    provider_values = source.get('provider_values')
    date_to_col = source.get('date_to_col')

    params = [start_date, end_date]
    time_filter_sql = sql.SQL("{date_col} >= %s AND {date_col} < %s").format(
        date_col=sql.Identifier(source['date_col'])
    )

    # Tryb ścisły: rekord musi mieścić się w danym miesiącu (np. 2026-02-01 .. 2026-02-28).
    if strict_month_window and date_to_col:
        end_of_month = _end_of_month_date(month_str)
        params = [start_date, end_of_month]
        time_filter_sql = sql.SQL(
            "{date_from} >= %s AND {date_to} <= %s"
        ).format(
            date_from=sql.Identifier(source['date_col']),
            date_to=sql.Identifier(date_to_col)
        )
    elif date_to_col:
        params = [end_date, start_date]
        time_filter_sql = sql.SQL(
            "{date_from} < %s AND ({date_to} IS NULL OR {date_to} >= %s)"
        ).format(
            date_from=sql.Identifier(source['date_col']),
            date_to=sql.Identifier(date_to_col)
        )

    provider_filter_sql = sql.SQL('')
    if provider_col and provider_values:
        placeholders = sql.SQL(', ').join([sql.Placeholder()] * len(provider_values))
        provider_filter_sql = sql.SQL(" AND UPPER(COALESCE({provider_col}::text, '')) IN ({vals})").format(
            provider_col=sql.Identifier(provider_col),
            vals=placeholders
        )
        params.extend([v.upper() for v in provider_values])

    query = sql.SQL(
        """
        SELECT
            {device_col} AS device_id,
            SUM(COALESCE({amount_col}, 0)) AS prowizja_zl,
            COUNT(*) AS liczba_rekordow
        FROM {schema}.{table}
        WHERE {time_filter}
          AND {device_filter}
          {provider_filter}
        GROUP BY {device_col}
        ORDER BY device_id
        """
    ).format(
        schema=sql.Identifier(source.get('schema', SCHEMA)),
        table=sql.Identifier(source['table']),
        device_col=sql.Identifier(source['device_col']),
        amount_col=sql.Identifier(source['amount_col']),
        time_filter=time_filter_sql,
        device_filter=_build_device_filter_sql(sql.Identifier(source['device_col'])),
        provider_filter=provider_filter_sql
    )

    cursor = conn.cursor()
    try:
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
    except psycopg2.Error as e:
        print(f"⚠ Błąd pobierania prowizji: {e}")
        rows = []
    finally:
        cursor.close()

    commission_data = {}
    for row in rows:
        device_id, prowizja, liczba_rekordow = row
        commission_data[int(device_id)] = {
            'prowizja_zl': float(prowizja or 0),
            'liczba_rekordow': int(liczba_rekordow)
        }

    return commission_data


def load_reference_provision_from_xls(xls_path):
    """
    Wczytuje plik referencyjny .xls (single sheet) i zwraca mapę prowizji per TVM_ID.
    Wymaga biblioteki pandas + xlrd. Jeśli brak bibliotek, zwraca None.
    """
    try:
        import pandas as pd
    except ImportError:
        print("⚠ Brak pandas - pomijam walidację pliku .xls")
        return None

    try:
        df = pd.read_excel(xls_path, sheet_name=0, engine='xlrd')
    except Exception as e:
        print(f"⚠ Nie udało się odczytać pliku .xls: {e}")
        return None

    normalized = {str(col).strip().upper(): col for col in df.columns}
    tvm_col = None
    amount_col = None

    for candidate in ['TVM_ID', 'TVMID', 'NR AUTOMATU', 'NR_AUTOMATU', 'DEVICE_ID']:
        if candidate in normalized:
            tvm_col = normalized[candidate]
            break

    for candidate in ['PROVISION_AMOUNT', 'PROWIZJA', 'PROWIZJA_ZL', 'KWOTA', 'AMOUNT']:
        if candidate in normalized:
            amount_col = normalized[candidate]
            break

    if tvm_col is None or amount_col is None:
        print("⚠ Nie znaleziono kolumn TVM_ID / PROVISION_AMOUNT w pliku referencyjnym")
        return None

    ref = {}
    for _, row in df[[tvm_col, amount_col]].dropna(subset=[tvm_col]).iterrows():
        try:
            device_id = int(row[tvm_col])
        except Exception:
            continue
        amount = row[amount_col]
        amount = 0.0 if amount is None else float(amount)
        ref[device_id] = amount

    return ref


def compare_commission_with_reference(commission_data, reference_map, tolerance=0.01):
    """
    Porównuje prowizję z SQL vs. plik referencyjny.
    """
    if reference_map is None:
        return None

    sql_map = {int(k): float(v.get('prowizja_zl', 0.0)) for k, v in commission_data.items()}
    sql_ids = set(sql_map.keys())
    ref_ids = set(reference_map.keys())

    missing_in_sql = sorted(ref_ids - sql_ids)
    extra_in_sql = sorted(sql_ids - ref_ids)
    mismatches = []

    for device_id in sorted(sql_ids & ref_ids):
        diff = abs(sql_map[device_id] - reference_map[device_id])
        if diff > tolerance:
            mismatches.append({
                'device_id': device_id,
                'sql_amount': sql_map[device_id],
                'xls_amount': reference_map[device_id],
                'difference': diff,
            })

    return {
        'missing_in_sql': missing_in_sql,
        'extra_in_sql': extra_in_sql,
        'mismatches': mismatches,
    }


def validate_provision_sql_vs_xls(conn, month_str, commission_source, xls_path, strict_month_window=True):
    """
    Waliduje prowizję: SQL (per TVM_ID) vs plik referencyjny PROVISION .xls (single sheet).
    Zwraca tuple: (comparison_dict_or_none, exit_code)
    """
    commission_data = get_monthly_commission(
        conn,
        month_str,
        source_config=commission_source,
        strict_month_window=strict_month_window,
    )
    reference_map = load_reference_provision_from_xls(xls_path)
    comparison = compare_commission_with_reference(commission_data, reference_map)

    if comparison is None:
        print("❌ Nie udało się porównać SQL vs XLS (sprawdź plik i biblioteki: pandas, xlrd)")
        return None, 2

    missing_count = len(comparison['missing_in_sql'])
    extra_count = len(comparison['extra_in_sql'])
    mismatch_count = len(comparison['mismatches'])

    print(f"✓ SQL rekordów (TVM_ID): {len(commission_data)}")
    reference_count = len(reference_map) if reference_map is not None else 0
    print(f"✓ XLS rekordów (TVM_ID): {reference_count}")
    print(f"  Brakujące w SQL: {missing_count}")
    print(f"  Nadmiarowe w SQL: {extra_count}")
    print(f"  Różnice kwot: {mismatch_count}")

    if missing_count:
        print(f"  Przykład brakujących TVM_ID w SQL: {comparison['missing_in_sql'][:10]}")
    if extra_count:
        print(f"  Przykład nadmiarowych TVM_ID w SQL: {comparison['extra_in_sql'][:10]}")
    if mismatch_count:
        print("  Przykłady różnic kwotowych:")
        for item in comparison['mismatches'][:10]:
            print(
                f"    TVM_ID={item['device_id']} SQL={item['sql_amount']:.2f} "
                f"XLS={item['xls_amount']:.2f} DIFF={item['difference']:.2f}"
            )

    exit_code = 0 if (missing_count == 0 and extra_count == 0 and mismatch_count == 0) else 1
    return comparison, exit_code


def get_monthly_revenue(
    conn,
    month_str,
    schema=SCHEMA,
    actioncodes=None,
    mctype=None,
    use_device_range_filter=True,
):
    """
    Pobiera obrót brutto per automat za dany miesiąc z tabeli <schema>.moneystats.
    Opcjonalnie filtruje po:
    - actioncodes: lista kodów ms_actioncode
    - mctype: wartość ms_mctype
    - use_device_range_filter: zakresy DEVICE_ID_RANGES
    month_str: 'YYYY-MM'
    Zwraca: dict {device_id: {'obrot_brutto_zl': float, 'liczba_transakcji': int}}
    """
    start_date, end_date = get_month_range(month_str)

    cursor = conn.cursor()

    params = [start_date, end_date]

    action_filter_sql = sql.SQL('')
    if actioncodes:
        placeholders = sql.SQL(', ').join([sql.Placeholder()] * len(actioncodes))
        action_filter_sql = sql.SQL(" AND ms_actioncode IN ({vals})").format(vals=placeholders)
        params.extend(actioncodes)

    mctype_filter_sql = sql.SQL('')
    if mctype is not None:
        mctype_filter_sql = sql.SQL(" AND ms_mctype = %s")
        params.append(mctype)

    device_filter_sql = sql.SQL('')
    if use_device_range_filter:
        device_filter_sql = sql.SQL(" AND {device_filter}").format(
            device_filter=_build_device_filter_sql(sql.Identifier('ms_deviceid'))
        )

    query = sql.SQL(
        """
        SELECT
            ms_deviceid AS device_id,
            SUM(COALESCE(ms_vtotal, 0)) AS obrot_brutto_zl,
            COUNT(*) AS liczba_transakcji
        FROM {schema}.moneystats
        WHERE ms_creadate >= %s
          AND ms_creadate < %s
          {action_filter}
          {mctype_filter}
          {device_filter}
        GROUP BY ms_deviceid
        ORDER BY device_id
        """
    ).format(
        schema=sql.Identifier(schema),
        action_filter=action_filter_sql,
        mctype_filter=mctype_filter_sql,
        device_filter=device_filter_sql,
    )

    cursor.execute(query, tuple(params))
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


def _normalize_carrier_code(carrier):
    """
    Normalizuje kod przewoźnika do postaci używanej w raporcie.
    """
    normalized = str(carrier or '').strip().upper()
    if normalized == 'KML':
        return 'KMŁ'
    return normalized


def _get_table_columns(conn, schema_name, table_name):
    """
    Zwraca listę kolumn tabeli lub pustą listę, gdy tabela nie istnieje.
    """
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT column_name
        FROM information_schema.columns
        WHERE table_schema = %s
          AND table_name = %s
        ORDER BY ordinal_position
        """,
        (schema_name, table_name),
    )
    rows = cursor.fetchall()
    cursor.close()
    return [r[0] for r in rows]


def _resolve_transactions_source(
    conn,
    schema_name,
    table_name,
    amount_col,
    date_col,
    device_col=None,
    carrier_col=None,
):
    """
    Rozwiązuje źródło transakcji z jednego schematu.
    Wykrywa kolumny urządzenia i przewoźnika lub używa jawnych nazw z CLI.
    Zwraca dict lub None, gdy brak poprawnego źródła.
    """
    columns = _get_table_columns(conn, schema_name, table_name)
    if not columns:
        return None

    lowered_map = {c.lower(): c for c in columns}
    if amount_col.lower() not in lowered_map or date_col.lower() not in lowered_map:
        return None

    if device_col:
        picked_device_col = lowered_map.get(device_col.lower())
    else:
        picked_device_col = _pick_first_matching_column(columns, TRANSACTIONS_DEVICE_COL_CANDIDATES)

    if picked_device_col is None:
        return None

    if carrier_col:
        picked_carrier_col = lowered_map.get(carrier_col.lower())
    else:
        picked_carrier_col = _pick_first_matching_column(columns, TRANSACTIONS_CARRIER_COL_CANDIDATES)

    return {
        'schema': schema_name,
        'table': table_name,
        'device_col': picked_device_col,
        'carrier_col': picked_carrier_col,
        'amount_col': lowered_map[amount_col.lower()],
        'date_col': lowered_map[date_col.lower()],
    }


def get_monthly_revenue_by_carrier(
    conn,
    month_str,
    schema=SCHEMA,
    carriers=None,
    table_name=TRANSACTIONS_SOURCE_CONFIG['table'],
    amount_col=TRANSACTIONS_SOURCE_CONFIG['amount_col'],
    date_col=TRANSACTIONS_SOURCE_CONFIG['date_col'],
    device_col=TRANSACTIONS_SOURCE_CONFIG['device_col'],
    carrier_col=None,
    use_device_range_filter=True,
):
    """
    Pobiera obrót brutto per automat i per przewoźnik z tabeli <schema>.transactions.
    Zwraca: dict {
        device_id: {
            'by_carrier': {
                carrier: {
                    'obrot_brutto_zl': float,
                    'liczba_transakcji': int
                }
            }
        }
    }
    """
    start_date, end_date = get_month_range(month_str)
    carriers = carriers or CARRIERS

    revenue_data = {}

    source = _resolve_transactions_source(
        conn,
        schema_name=schema,
        table_name=table_name,
        amount_col=amount_col,
        date_col=date_col,
        device_col=device_col,
        carrier_col=carrier_col,
    )
    if source is None:
        print(
            f"⚠ Pomijam tryb per przewoźnik: brak źródła {schema}.{table_name} "
            f"lub wymaganych kolumn (date/amount/device)"
        )
        return {}

    if source.get('carrier_col') is None:
        print(
            "⚠ Nie znaleziono kolumny przewoźnika w transactions. "
            "Pomijam tryb per przewoźnik, aby uniknąć dublowania danych."
        )
        return {}

    for carrier in carriers:
        carrier_code = _normalize_carrier_code(carrier)

        params = [start_date, end_date, carrier_code]
        device_filter_sql = sql.SQL('')
        if use_device_range_filter:
            device_filter_sql = sql.SQL(" AND {device_filter}").format(
                device_filter=_build_device_filter_sql(sql.Identifier(source['device_col']))
            )
        carrier_filter_sql = sql.SQL(
            " AND UPPER(COALESCE({carrier_col}::text, '')) = %s"
        ).format(
            carrier_col=sql.Identifier(source['carrier_col'])
        )

        query = sql.SQL(
            """
            SELECT
                {device_col} AS device_id,
                SUM(COALESCE({amount_col}, 0)) AS obrot_brutto_zl,
                COUNT(*) AS liczba_transakcji
            FROM {schema}.{table}
            WHERE {date_col} >= %s
              AND {date_col} < %s
              {device_filter}
                            {carrier_filter}
            GROUP BY {device_col}
            ORDER BY device_id
            """
        ).format(
            schema=sql.Identifier(source['schema']),
            table=sql.Identifier(source['table']),
            device_col=sql.Identifier(source['device_col']),
            amount_col=sql.Identifier(source['amount_col']),
            date_col=sql.Identifier(source['date_col']),
            device_filter=device_filter_sql,
            carrier_filter=carrier_filter_sql,
        )

        cursor = conn.cursor()
        try:
            cursor.execute(query, tuple(params))
            rows = cursor.fetchall()
            print(
                f"✓ Obrót {carrier_code}: {source['schema']}.{source['table']} "
                f"({source['device_col']}, {source['carrier_col']}, {source['amount_col']}, {source['date_col']}) -> {len(rows)} automatów"
            )
        except psycopg2.Error as e:
            print(f"⚠ Błąd pobierania obrotu dla {carrier_code}: {e}")
            rows = []
        finally:
            cursor.close()

        for row in rows:
            device_id, obrot, liczba = row
            try:
                device_id_int = int(device_id)
            except (TypeError, ValueError):
                continue

            revenue_data.setdefault(device_id_int, {'by_carrier': {}})
            revenue_data[device_id_int]['by_carrier'][carrier_code] = {
                'obrot_brutto_zl': float(obrot or 0.0),
                'liczba_transakcji': int(liczba or 0),
            }

    return revenue_data


# === EKSPORT DO EXCEL (KOLUMNY 1-2, RESZTA PLACEHOLDER) ===

def export_to_excel_PL(dictionary_comparison, revenue_data, month_str, filename, commission_data=None):
    """
    Eksportuje raport P&L do Excela.
    Kolumny:
    1-INFO: nr_automatu, lokalizacja, status_zmiany
    2-OBRÓT: obrot_brutto_zl, liczba_transakcji
    3-PROWIZJA: prowizja_zl
    5,7,8: placeholder (do ETAPU 2 i 3)
    """
    if commission_data is None:
        commission_data = {}

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    filepath = OUTPUT_DIR / filename
    
    wb = Workbook()
    ws = wb.active
    ws.title = f"P&L {month_str}"
    
    # === NAGŁÓWKI ===
    headers = [
        # 1-INFO
        'Nr automatu', 'Lokalizacja', 'Status', 'Przewoźnik',
        # 2-OBRÓT
        'Obrót brutto (zł)', 'Liczba transakcji',
        # 3-PROWIZJA
        'Prowizja (zł)',
        # 5-KOSZTY (placeholder)
        'Koszty (zł)',
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
            lokalizacja = extract_city_name(data.get('description', ''))
        else:
            lokalizacja = ''
            status = 'brak w dict'
        
        # Obsługa nowego modelu: revenue per (automat, przewoźnik).
        rev_by_carrier = {}
        if device_id in revenue_data:
            rev_entry = revenue_data[device_id]
            if isinstance(rev_entry, dict) and isinstance(rev_entry.get('by_carrier'), dict):
                rev_by_carrier = rev_entry['by_carrier']
            elif isinstance(rev_entry, dict):
                rev_by_carrier = {'': rev_entry}

        comm_by_carrier = {}
        if device_id in commission_data:
            comm_entry = commission_data[device_id]
            if isinstance(comm_entry, dict) and isinstance(comm_entry.get('by_carrier'), dict):
                comm_by_carrier = comm_entry['by_carrier']
            elif isinstance(comm_entry, dict):
                comm_by_carrier = {'': comm_entry}

        carriers_for_device = sorted(set(rev_by_carrier.keys()) | set(comm_by_carrier.keys()))
        if not carriers_for_device:
            carriers_for_device = ['']

        for carrier in carriers_for_device:
            # 2-OBRÓT
            rev = rev_by_carrier.get(carrier)
            if isinstance(rev, dict):
                obrot = rev.get('obrot_brutto_zl', 0.0)
                liczba = rev.get('liczba_transakcji', 0)
            else:
                obrot = 0.0
                liczba = 0

            # 3-PROWIZJA
            comm = comm_by_carrier.get(carrier)
            if isinstance(comm, dict):
                prowizja = comm.get('prowizja_zl', 0.0)
            else:
                prowizja = ''

            # Placeholder dla kolumn 5,7,8
            koszty = ''
            suma_roczna = ''
            uwagi = ''

            # Wpisz dane do wiersza
            ws.cell(row=row_num, column=1).value = device_id
            ws.cell(row=row_num, column=2).value = lokalizacja
            ws.cell(row=row_num, column=3).value = status
            ws.cell(row=row_num, column=4).value = carrier
            ws.cell(row=row_num, column=5).value = obrot
            ws.cell(row=row_num, column=5).number_format = '#,##0.00'
            ws.cell(row=row_num, column=6).value = int(liczba)
            ws.cell(row=row_num, column=7).value = prowizja
            if isinstance(prowizja, (int, float)):
                ws.cell(row=row_num, column=7).number_format = '#,##0.00'
            ws.cell(row=row_num, column=8).value = koszty
            ws.cell(row=row_num, column=9).value = suma_roczna
            ws.cell(row=row_num, column=10).value = uwagi

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
        '--schema',
        type=str,
        default=SCHEMA,
        help='Schemat źródłowy dla dictionary i moneystats (domyślnie: AV)'
    )
    parser.add_argument(
        '--obrot-actioncodes',
        nargs='*',
        type=int,
        default=None,
        help='Filtr obrotu: lista ms_actioncode (np. --obrot-actioncodes 1 2)'
    )
    parser.add_argument(
        '--obrot-mctype',
        type=int,
        default=None,
        help='Filtr obrotu: ms_mctype (np. --obrot-mctype 4)'
    )
    parser.add_argument(
        '--obrot-no-device-range-filter',
        action='store_true',
        help='Wyłącza filtr zakresów urządzeń dla obrotu (1101-1141, 1201-1299)'
    )
    parser.add_argument(
        '--obrot-source-mode',
        choices=['carrier-transactions', 'moneystats'],
        default='carrier-transactions',
        help='Źródło obrotu: per przewoźnik z <schema>.transactions lub klasycznie z <schema>.moneystats'
    )
    parser.add_argument(
        '--obrot-carriers',
        nargs='*',
        default=CARRIERS,
        help='Lista przewoźników do pobrania (np. IC PR KD ARP KW KS LKA SKM KMŁ)'
    )
    parser.add_argument(
        '--obrot-transactions-table',
        type=str,
        default=TRANSACTIONS_SOURCE_CONFIG['table'],
        help='Nazwa tabeli transakcji per przewoźnik (domyślnie: transactions)'
    )
    parser.add_argument(
        '--obrot-transactions-amount-col',
        type=str,
        default=TRANSACTIONS_SOURCE_CONFIG['amount_col'],
        help='Kolumna kwoty w tabeli transakcji (domyślnie: fin_ptu_kwota)'
    )
    parser.add_argument(
        '--obrot-transactions-date-col',
        type=str,
        default=TRANSACTIONS_SOURCE_CONFIG['date_col'],
        help='Kolumna daty w tabeli transakcji (domyślnie: fin_data_sp)'
    )
    parser.add_argument(
        '--obrot-transactions-device-col',
        type=str,
        default=TRANSACTIONS_SOURCE_CONFIG['device_col'],
        help='Kolumna numeru automatu (opcjonalnie, domyślnie auto-detekcja)'
    )
    parser.add_argument(
        '--obrot-transactions-carrier-col',
        type=str,
        default=None,
        help='Kolumna przewoźnika w tabeli transactions (opcjonalnie, domyślnie auto-detekcja)'
    )
    parser.add_argument(
        '--miesiac',
        type=str,
        help='Miesiąc raportu w formacie YYYY-MM (domyślnie: poprzedni miesiąc)',
        default=None
    )
    parser.add_argument('--prowizja-schema', type=str, default=COMMISSION_SOURCE_CONFIG['schema'], help='Schema tabeli prowizji (domyślnie: public)')
    parser.add_argument('--prowizja-table', type=str, default=COMMISSION_SOURCE_CONFIG['table'], help='Nazwa tabeli prowizji (domyślnie: provision)')
    parser.add_argument('--prowizja-device-col', type=str, default=COMMISSION_SOURCE_CONFIG['device_col'], help='Kolumna z numerem automatu (np. TVM_ID)')
    parser.add_argument('--prowizja-amount-col', type=str, default=COMMISSION_SOURCE_CONFIG['amount_col'], help='Kolumna z kwotą prowizji (np. provision_amount)')
    parser.add_argument('--prowizja-date-col', type=str, default=COMMISSION_SOURCE_CONFIG['date_col'], help='Kolumna daty początkowej')
    parser.add_argument('--prowizja-date-to-col', type=str, default=COMMISSION_SOURCE_CONFIG['date_to_col'], help='Kolumna daty końca obowiązywania (opcjonalnie)')
    parser.add_argument('--prowizja-provider-col', type=str, default=None, help='Kolumna z operatorem płatności (np. ELAVON)')
    parser.add_argument(
        '--prowizja-provider-values',
        nargs='*',
        default=None,
        help='Dozwolone wartości operatora płatności (np. ELAVON INTERCHANGE)'
    )
    parser.add_argument(
        '--strict-prowizja-month-window',
        action='store_true',
        help='Ścisły zakres dla prowizji: date_from >= 1 dzień miesiąca i date_to <= ostatni dzień miesiąca'
    )
    parser.add_argument(
        '--no-strict-prowizja-month-window',
        action='store_false',
        dest='strict_prowizja_month_window',
        help='Tryb nakładających się okresów (date_from < koniec_miesiąca i date_to >= początek_miesiąca)'
    )
    parser.add_argument(
        '--reference-provision-xls',
        type=str,
        default=None,
        help='Ścieżka do pliku .xls z raportu PROVISION (single sheet) do porównania'
    )
    parser.add_argument(
        '--output-naming',
        choices=['default', 'monitor-style'],
        default='default',
        help='Tryb nazewnictwa pliku wyjściowego'
    )
    parser.add_argument(
        '--validate-provision-only',
        action='store_true',
        help='Uruchamia tylko walidację prowizji SQL vs plik PROVISION .xls (bez obrotu i bez eksportu)'
    )
    parser.set_defaults(strict_prowizja_month_window=True)
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
    
    commission_source = {
        'schema': args.prowizja_schema,
        'table': args.prowizja_table,
        'device_col': args.prowizja_device_col,
        'amount_col': args.prowizja_amount_col,
        'date_col': args.prowizja_date_col,
        'date_to_col': args.prowizja_date_to_col,
        'provider_col': args.prowizja_provider_col,
        'provider_values': args.prowizja_provider_values
    }

    source_schema = args.schema

    print(f"\n{'='*60}")
    print(f"  Miesięczne zestawienie TVM P&L - {month_str}")
    print(f"  Schemat danych: {source_schema}")
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

    if args.validate_provision_only:
        if not args.reference_provision_xls:
            print("❌ Dla trybu --validate-provision-only podaj --reference-provision-xls")
            conn.close()
            return

        print("\n[VALIDATION] SQL vs PROVISION XLS...")
        _, validation_exit_code = validate_provision_sql_vs_xls(
            conn,
            month_str,
            commission_source,
            args.reference_provision_xls,
            strict_month_window=args.strict_prowizja_month_window,
        )
        conn.close()
        if validation_exit_code == 0:
            print("\n✓ Walidacja zakończona: pełna zgodność SQL vs XLS")
        elif validation_exit_code == 1:
            print("\n⚠ Walidacja zakończona: wykryto różnice SQL vs XLS")
        return
    
    # === KROK 1: Słownik automatów (ARP.dictionary) ===
    print(f"\n[1/5] Pobieranie słownika automatów ({source_schema}.dictionary)...")
    current_snapshot = get_dictionary_snapshot(conn, schema=source_schema)
    print(f"✓ Pobrano {len(current_snapshot)} automatów ze słownika")
    
    # Porównaj z poprzednim miesiącem
    print("\n[2/5] Porównanie z poprzednim miesiącem...")
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
    revenue_data = {}
    if args.obrot_source_mode == 'carrier-transactions':
        print(
            f"\n[3/5] Pobieranie obrotu za {month_str} "
            f"(per przewoźnik: {source_schema}.{args.obrot_transactions_table})..."
        )
        if args.obrot_no_device_range_filter:
            print("  Filtr zakresów urządzeń: WYŁĄCZONY")
        print(f"  Przewoźnicy: {', '.join(args.obrot_carriers)}")

        revenue_data = get_monthly_revenue_by_carrier(
            conn,
            month_str,
            schema=source_schema,
            carriers=args.obrot_carriers,
            table_name=args.obrot_transactions_table,
            amount_col=args.obrot_transactions_amount_col,
            date_col=args.obrot_transactions_date_col,
            device_col=args.obrot_transactions_device_col,
            carrier_col=args.obrot_transactions_carrier_col,
            use_device_range_filter=not args.obrot_no_device_range_filter,
        )

        total_revenue = sum(
            carrier_values.get('obrot_brutto_zl', 0.0)
            for device_values in revenue_data.values()
            for carrier_values in device_values.get('by_carrier', {}).values()
        )
        total_transactions = sum(
            int(carrier_values.get('liczba_transakcji', 0) or 0)
            for device_values in revenue_data.values()
            for carrier_values in device_values.get('by_carrier', {}).values()
        )
        print(f"✓ Pobrano dane obrotu dla {len(revenue_data)} automatów (model per przewoźnik)")
    else:
        print(f"\n[3/5] Pobieranie obrotu za {month_str} ({source_schema}.moneystats)...")
        if args.obrot_actioncodes:
            print(f"  Filtr ms_actioncode: {args.obrot_actioncodes}")
        if args.obrot_mctype is not None:
            print(f"  Filtr ms_mctype: {args.obrot_mctype}")
        if args.obrot_no_device_range_filter:
            print("  Filtr zakresów urządzeń: WYŁĄCZONY")

        revenue_data = get_monthly_revenue(
            conn,
            month_str,
            schema=source_schema,
            actioncodes=args.obrot_actioncodes,
            mctype=args.obrot_mctype,
            use_device_range_filter=not args.obrot_no_device_range_filter,
        )
        print(f"✓ Pobrano dane obrotu dla {len(revenue_data)} automatów")
        total_revenue = sum(v['obrot_brutto_zl'] for v in revenue_data.values())
        total_transactions = sum(v['liczba_transakcji'] for v in revenue_data.values())

    print(f"  Obrót brutto: {total_revenue:,.2f} zł")
    print(f"  Liczba transakcji: {total_transactions:,}")

    # === KROK 3: Prowizja miesięczna ===
    print(f"\n[4/5] Pobieranie prowizji za {month_str}...")
    commission_data = get_monthly_commission(
        conn,
        month_str,
        source_config=commission_source,
        strict_month_window=args.strict_prowizja_month_window,
    )
    print(f"✓ Pobrano dane prowizji dla {len(commission_data)} automatów")
    total_commission = sum(v['prowizja_zl'] for v in commission_data.values())
    print(f"  Prowizja łączna: {total_commission:,.2f} zł")

    if args.reference_provision_xls:
        print("  🔎 Porównanie prowizji SQL vs plik referencyjny .xls...")
        reference_map = load_reference_provision_from_xls(args.reference_provision_xls)
        comparison = compare_commission_with_reference(commission_data, reference_map)
        if comparison is None:
            print("  ⚠ Pominięto porównanie (nie udało się odczytać pliku .xls)")
        else:
            print(f"  Brakujące w SQL: {len(comparison['missing_in_sql'])}")
            print(f"  Nadmiarowe w SQL: {len(comparison['extra_in_sql'])}")
            print(f"  Różnice kwot: {len(comparison['mismatches'])}")
    
    conn.close()
    
    # === KROK 4: Eksport do Excel ===
    print(f"\n[5/5] Eksport do Excel...")
    filename = build_output_filename(month_str, naming_mode=args.output_naming)
    export_to_excel_PL(dictionary_comparison, revenue_data, month_str, filename, commission_data=commission_data)
    
    print(f"\n{'='*60}")
    print(f"  ✓ RAPORT ZAKOŃCZONY POMYŚLNIE")
    print(f"  Kolumny zaimplementowane: 1-INFO, 2-OBRÓT, 3-PROWIZJA")
    print(f"  Do zrobienia (ETAP 2): kolumny 4-6")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()