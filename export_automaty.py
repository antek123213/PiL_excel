# export_automaty.py - Miesięczne zestawienie TVM P&L (ETAP 1: kolumny 1-INFO, 2-OBRÓT, 3-PROWIZJA)
import psycopg2
from psycopg2 import sql
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import json
import argparse
import time
from pathlib import Path
import re
import calendar
import unicodedata
import os
import sys


def _silent_print(*args, **kwargs):
    return None


if os.environ.get('PIL_SILENT', '').strip() == '1':
    print = _silent_print

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
DEFAULT_COSTS_FILE = Path(__file__).parent / 'P&L_2025.11.25_koszty_ROP 2025.xlsx'
DEFAULT_RENT_FILE = Path(__file__).parent / 'Najem powierzchni 2025-2026.xlsx'
DEFAULT_AMORTYZACJA_FILE = Path(__file__).parent / 'Amortyzacja miesieczna Automaty.xlsx'
DEFAULT_PROWIZJE_FILE = Path(__file__).parent / 'Prowizje_AB.xlsx'
DEFAULT_AMORTYZACJA_SHEET = 'bb8'
DEVICE_ID_RANGES = (
    (1101, 1141),
    (1201, 1299),
)

CARRIERS = ['ARP', 'IC', 'KD', 'KML', 'KW', 'LKA', 'PR', 'SKM']
DISALLOWED_CARRIERS = {'KM'}
TRANSACTIONS_SOURCE_CONFIG = {
    'table': 'transactions',
    'amount_col': 'fin_nalezn',
    'date_col': 'fin_data_sp',
    'device_col': 'tvm_tvm_id',
    'payment_method_col': 'fin_spos_opl',
}
TRANSACTIONS_TABLE_OVERRIDES = {
    'IC': 'transactionsns',
}
TRANSACTIONS_DEVICE_COL_OVERRIDES = {
    # User-required precedence: tvm_aitomatnum first, then tvm_automatnum.
    'IC': ['tvm_aitomatnum', 'tvm_automatnum'],
}
TRANSACTIONS_AMOUNT_COL_OVERRIDES = {
}
TRANSACTIONS_PAYMENT_METHOD_COL_OVERRIDES = {
}
TRANSACTIONS_PAYMENT_METHOD_COL_CANDIDATES = [
    'fin_spos_opl',
    'spos_opl',
    'payment_method',
    'payment_method_code',
]
PAYMENT_METHOD_CODE_TO_KEY = {
    '1': 'gotowka',
    '2': 'karta',
    '6': 'blik',
}
PAYMENT_METHOD_KEYS = ('gotowka', 'karta', 'blik')
PAYMENT_METHOD_LABELS = {
    'gotowka': 'Gotówka',
    'karta': 'Karta',
    'blik': 'BLIK',
}
IT_CARD_RATE_BY_CARRIER = {
    'KD': 0.0154,
    'KML': 0.0135,
    'LKA': 0.0154,
    'ARP': 0.0154,
    'PR': 0.0154,
}
DEFAULT_IT_CARD_RATE = 0.0154
TVM_COST_KEYS = (
    'czynsz',
    'prad',
    'elavon',
    'poczta_polska',
    'amortyzacja',
    'utrzymanie_oprogramowania',
    'papier',
    'transmisja_danych',
    'serwis',
    'it_card',
    'ubezpieczenie',
)
TVM_COST_LABELS = {
    'czynsz': 'czynsz',
    'prad': 'prad',
    'elavon': 'elavon',
    'poczta_polska': 'poczta polska',
    'amortyzacja': 'amortyzacja',
    'utrzymanie_oprogramowania': 'utrzymanie oprogramowania',
    'papier': 'papier',
    'transmisja_danych': 'transmisja danych',
    'serwis': 'serwis',
    'it_card': 'IT card',
    'ubezpieczenie': 'ubezpieczenie',
}
OTHER_COST_KEYS = ('non_tvm', 'project_variable_costs', 'zdankiewicz', 'oh')
OTHER_COST_LABELS = {
    'non_tvm': 'NON TVM',
    'project_variable_costs': 'Project Variable Costs',
    'zdankiewicz': 'Zdankiewicz najklejki',
    'oh': 'OH',
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


def _normalize_tvm_location_text(value):
    """
    Normalizuje pełny tekst lokalizacji z pól TVM.
    """
    if value is None:
        return ''
    normalized = str(value).strip()
    if not normalized:
        return ''
    if 'opis' in normalized.lower():
        return ''
    return normalized


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


def _extract_automat_type_suffix(groupid_value):
    """
    Wyciąga suffix typu automatu z pola groupid, np. AUTABB -> BB.
    """
    normalized = str(groupid_value or '').strip().upper()
    if not normalized.startswith('AUTA'):
        return ''
    return normalized[4:]


def get_automat_type_by_device_from_av(
    conn,
    device_ids,
    schema_name='AV',
    table_name='dictionary',
    value_col='value',
    groupid_col='groupid',
):
    """
    Pobiera typ automatu z AV.dictionary po kolumnie `value` i zwraca mapę:
    {device_id: 'BB' | ...}.
    """
    normalized_ids = sorted({int(device_id) for device_id in device_ids})
    if not normalized_ids:
        return {}

    columns = _get_table_columns(conn, schema_name, table_name)
    if not columns:
        print(f"⚠ Brak tabeli typu automatu: {schema_name}.{table_name}")
        return {}

    lowered = {c.lower(): c for c in columns}
    resolved_value_col = lowered.get(str(value_col).lower())
    resolved_groupid_col = lowered.get(str(groupid_col).lower())
    if resolved_value_col is None or resolved_groupid_col is None:
        print(
            f"⚠ Brak kolumn typu automatu w {schema_name}.{table_name}: "
            f"{value_col}, {groupid_col}"
        )
        return {}

    search_values = [str(device_id) for device_id in normalized_ids]
    cursor = conn.cursor()
    query = sql.SQL(
        """
        SELECT {value_col}, {groupid_col}
        FROM {schema}.{table}
        WHERE {value_col}::text = ANY(%s)
        """
    ).format(
        value_col=sql.Identifier(resolved_value_col),
        groupid_col=sql.Identifier(resolved_groupid_col),
        schema=sql.Identifier(schema_name),
        table=sql.Identifier(table_name),
    )
    cursor.execute(query, (search_values,))
    rows = cursor.fetchall()
    cursor.close()

    type_map = {}
    for raw_value, groupid in rows:
        device_id = _parse_device_id(raw_value)
        if device_id is None:
            continue
        type_map[device_id] = _extract_automat_type_suffix(groupid)

    return type_map


def _resolve_dictionary_location_source(conn, schema_name, table_name='dictionary'):
    """
    Rozpoznaje kolumny słownika dla lokalizacji (id, tvm_tvmlocation1, tvm_tvmlocation2, description).
    Zwraca dict lub None, gdy tabela/kolumny nie są dostępne.
    """
    columns = _get_table_columns(conn, schema_name, table_name)
    if not columns:
        return None

    lowered_map = {c.lower(): c for c in columns}
    device_col = lowered_map.get('id')
    tvm_location1_col = lowered_map.get('tvm_tvmlocation1')
    tvm_location2_col = lowered_map.get('tvm_tvmlocation2')
    description_col = lowered_map.get('description')

    if device_col is None:
        return None
    if tvm_location1_col is None and tvm_location2_col is None and description_col is None:
        return None

    return {
        'schema': schema_name,
        'table': table_name,
        'device_col': device_col,
        'tvm_location1_col': tvm_location1_col,
        'tvm_location2_col': tvm_location2_col,
        'description_col': description_col,
    }


def _pick_location_value(tvm_location1, tvm_location2, description):
    """
    Wybiera lokalizację z priorytetem:
    tvm_tvmlocation1 -> tvm_tvmlocation2 -> description.

    Dla pól TVM zachowuje pełny adres.
    Description jest używane tylko, gdy oba pola TVM są puste.
    Zwraca pusty string, gdy brak sensownej wartości.
    """
    normalized_1 = _normalize_tvm_location_text(tvm_location1)
    normalized_2 = _normalize_tvm_location_text(tvm_location2)

    if normalized_1 and normalized_2:
        return f"{normalized_1}, {normalized_2}"
    if normalized_1:
        return normalized_1
    if normalized_2:
        return normalized_2

    return extract_city_name(description)


def get_locations_by_carrier(
    conn,
    device_ids,
    carriers=None,
    dictionary_table='dictionary',
):
    """
    Uzupełnia lokalizacje automatów sekwencyjnie po przewoźnikach.
    Priorytet pól: tvm_tvmlocation1 -> tvm_tvmlocation2 -> description.
    Dla automatu, który już ma lokalizację, kolejne schematy są pomijane.

    Zwraca dict {device_id: 'lokalizacja'}.
    """
    carriers = carriers or CARRIERS
    missing_device_ids = {int(device_id) for device_id in device_ids}
    location_by_device = {}

    for carrier in carriers:
        if not missing_device_ids:
            break

        carrier_code = _normalize_carrier_code(carrier)
        source = _resolve_dictionary_location_source(
            conn,
            schema_name=carrier_code,
            table_name=dictionary_table,
        )
        if source is None:
            print(
                f"⚠ Pomijam lokalizacje dla {carrier} (schemat {carrier_code}): "
                f"brak źródła {carrier_code}.{dictionary_table} lub wymaganych kolumn"
            )
            continue

        selected_columns = [
            sql.Identifier(source['device_col']),
            sql.Identifier(source['tvm_location1_col']) if source['tvm_location1_col'] else sql.SQL('NULL'),
            sql.Identifier(source['tvm_location2_col']) if source['tvm_location2_col'] else sql.SQL('NULL'),
            sql.Identifier(source['description_col']) if source['description_col'] else sql.SQL('NULL'),
        ]

        query = sql.SQL(
            """
                        SELECT {device_col}, {tvm_location1_col}, {tvm_location2_col}, {description_col}
            FROM {schema}.{table}
            WHERE {device_col}::text ~ '^[0-9]+$'
              AND {device_filter}
            ORDER BY {device_col}
            """
        ).format(
            device_col=selected_columns[0],
                        tvm_location1_col=selected_columns[1],
                        tvm_location2_col=selected_columns[2],
                        description_col=selected_columns[3],
            schema=sql.Identifier(source['schema']),
            table=sql.Identifier(source['table']),
            device_filter=_build_device_filter_sql(
                sql.SQL("{device_col}::bigint").format(device_col=selected_columns[0])
            ),
        )

        cursor = conn.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        cursor.close()

        filled_for_carrier = 0
        for row in rows:
            device_id, tvm_location1, tvm_location2, description = row
            normalized_device_id = int(device_id)
            if normalized_device_id not in missing_device_ids:
                continue

            picked_location = _pick_location_value(tvm_location1, tvm_location2, description)
            if not picked_location:
                continue

            location_by_device[normalized_device_id] = picked_location
            missing_device_ids.remove(normalized_device_id)
            filled_for_carrier += 1

        print(
            f"✓ Lokalizacje: {carrier_code} -> +{filled_for_carrier}, "
            f"braki po kroku: {len(missing_device_ids)}"
        )

    return location_by_device


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


def get_month_range_closed(month_str):
    """
    Zwraca domknięty zakres dat dla miesiąca raportowego:
    od 1 dnia 00:00:00 do ostatniego dnia 23:59:59.
    """
    start_date, _ = get_month_range(month_str)
    year, month = map(int, month_str.split('-'))
    last_day = calendar.monthrange(year, month)[1]
    end_date = datetime(year, month, last_day, 23, 59, 59)
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


def _has_nonempty_value(value):
    """
    Sprawdza, czy wartość jest niepusta z perspektywy odczytu z Excela.
    """
    if value is None:
        return False
    return bool(str(value).strip())


def _as_float(value, as_percent=False):
    """
    Konwertuje wartość z Excela na float.

    Obsługuje m.in. polski zapis dziesiętny, separatory tysięcy,
    symbole walut i zapis procentowy.
    Dla as_percent=True zwraca ułamek (np. 2% -> 0.02, 2 -> 0.02).
    """
    if value is None or isinstance(value, bool):
        return 0.0

    if isinstance(value, (int, float)):
        numeric = float(value)
        if not as_percent:
            return numeric
        if abs(numeric) <= 1.0:
            return numeric
        return numeric / 100.0

    text = str(value).strip()
    if not text:
        return 0.0

    text = (
        text
        .replace('\xa0', ' ')
        .replace('\u202f', ' ')
        .replace('\u2009', ' ')
        .replace('−', '-')
        .replace('–', '-')
        .replace('—', '-')
    )

    is_parentheses_negative = text.startswith('(') and text.endswith(')')
    if is_parentheses_negative:
        text = text[1:-1].strip()

    has_percent = '%' in text
    text = re.sub(r'(?i)\b(PLN|ZL|ZŁ|EUR|USD)\b', '', text)
    text = text.replace('%', '')
    text = text.replace(' ', '')
    text = text.replace("'", '')

    text = re.sub(r'[^0-9,\.\-]', '', text)
    if not text or text in ('-', '--'):
        return 0.0

    if text.startswith('-'):
        text = '-' + text[1:].replace('-', '')
    else:
        text = text.replace('-', '')

    if ',' in text and '.' in text:
        if text.rfind(',') > text.rfind('.'):
            text = text.replace('.', '')
            text = text.replace(',', '.')
        else:
            text = text.replace(',', '')
    elif ',' in text:
        if text.count(',') == 1:
            text = text.replace(',', '.')
        else:
            parts = text.split(',')
            text = ''.join(parts[:-1]) + '.' + parts[-1]
    elif text.count('.') > 1:
        parts = text.split('.')
        text = ''.join(parts[:-1]) + '.' + parts[-1]

    try:
        numeric = float(text)
    except (TypeError, ValueError):
        return 0.0

    if is_parentheses_negative and numeric > 0:
        numeric = -numeric

    if not as_percent:
        return numeric

    if has_percent:
        return numeric / 100.0
    if abs(numeric) <= 1.0:
        return numeric
    return numeric / 100.0


def _matches_month(value, month_str):
    """
    Sprawdza czy wartość daty należy do miesiąca raportowego YYYY-MM.
    """
    if value is None:
        return False
    if hasattr(value, 'year') and hasattr(value, 'month'):
        return f"{value.year:04d}-{value.month:02d}" == month_str
    return False


def _extract_monthly_value_in_section(ws, section_label, month_str, month_col=2, value_col=3):
    """
    Odczytuje miesięczną wartość z bloku sekcji oznaczonej etykietą w kolumnie B.
    """
    section_row = None
    normalized_label = str(section_label or '').strip().upper()
    for r in range(1, ws.max_row + 1):
        cell_val = ws.cell(r, month_col).value
        if isinstance(cell_val, str) and str(cell_val).strip().upper() == normalized_label:
            section_row = r
            break

    if section_row is None:
        return 0.0

    value = 0.0
    for r in range(section_row + 1, ws.max_row + 1):
        marker = ws.cell(r, month_col).value
        if isinstance(marker, str) and marker.strip():
            break
        if _matches_month(marker, month_str):
            value += _as_float(ws.cell(r, value_col).value)
    return value


def _find_section_row(ws, section_label, label_col=2):
    """
    Zwraca numer wiersza sekcji opisanej etykietą w zadanej kolumnie.
    """
    normalized_label = str(section_label or '').strip().upper()
    for r in range(1, ws.max_row + 1):
        cell_val = ws.cell(r, label_col).value
        if isinstance(cell_val, str) and str(cell_val).strip().upper() == normalized_label:
            return r
    return None


def _extract_monthly_value_from_columns(ws, month_str, month_col, value_col, start_row=1, end_row=None):
    """
    Sumuje wartości dla miesiąca w zadanych kolumnach (układ tabelaryczny bez sekcji).
    """
    if end_row is None:
        end_row = ws.max_row

    total = 0.0
    for r in range(max(1, int(start_row)), min(ws.max_row, int(end_row)) + 1):
        month_cell = ws.cell(r, month_col).value
        if _matches_month(month_cell, month_str):
            total += _as_float(ws.cell(r, value_col).value)
    return total


def _get_section_end_row(ws, section_row, section_label_col=2):
    """
    Szuka końca sekcji: pierwszy kolejny wiersz z nową etykietą sekcji.
    """
    if section_row is None:
        return ws.max_row

    for r in range(section_row + 1, ws.max_row + 1):
        marker = ws.cell(r, section_label_col).value
        if isinstance(marker, str) and marker.strip():
            return r - 1
    return ws.max_row


def _extract_monthly_value_from_header_table(ws, month_str, header_label, month_col, value_col, header_col=None):
    """
    Odczytuje wartość dla miesiąca z tabeli identyfikowanej nagłówkiem (np. Non TVMs/Project).
    Zakłada, że dane tabeli są w kolejnych wierszach pod nagłówkiem i kończą się przy pierwszej
    pustej komórce miesiąca po rozpoczęciu bloku dat.
    """
    header_row = _find_section_row(ws, header_label, label_col=header_col or value_col)
    if header_row is None:
        return 0.0

    total = 0.0
    seen_date_rows = False
    for r in range(header_row + 1, ws.max_row + 1):
        month_cell = ws.cell(r, month_col).value

        if _matches_month(month_cell, month_str):
            total += _as_float(ws.cell(r, value_col).value)
            seen_date_rows = True
            continue

        if hasattr(month_cell, 'year') and hasattr(month_cell, 'month'):
            seen_date_rows = True
            continue

        if month_cell is None and seen_date_rows:
            break

        if isinstance(month_cell, str) and month_cell.strip():
            break

    return total


def _normalize_header_text(value):
    """
    Normalizuje tekst nagłówka do porównań (usuwa spacje, kropki i znaki specjalne).
    """
    if value is None:
        return ''
    return re.sub(r'[^A-Z0-9]+', '', str(value).strip().upper())


def _normalize_text_for_match(value):
    """
    Normalizuje tekst do porównań fuzzy (lokalizacje, nazwy) usuwając separatory.
    """
    if value is None:
        return ''
    return re.sub(r'[^A-Z0-9]+', '', str(value).strip().upper())


def _build_month_year_labels(month_str):
    """
    Zwraca etykiety miesiąca w formatach użytecznych do wyszukiwania nagłówków.
    """
    dt = datetime.strptime(month_str, '%Y-%m')
    return dt.strftime('%m/%Y'), dt.strftime('%m%Y')


def _parse_device_id(value):
    """
    Konwertuje numer automatu do int; zwraca None dla wartości niejednoznacznych.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return None

    text = str(value).strip().replace(',', '.')
    if not text:
        return None
    if re.fullmatch(r'\d+(\.0+)?', text):
        return int(float(text))
    return None


def _as_date(value):
    """
    Konwertuje wartość do date; zwraca None gdy brak poprawnej daty.
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if hasattr(value, 'date'):
        return value.date()
    return None


def _month_date_bounds(month_str):
    """
    Zwraca zakres dat miesiąca raportowego jako date (start, end).
    """
    start_dt, end_dt = get_month_range_closed(month_str)
    return start_dt.date(), end_dt.date()


def _normalize_excel_commission_header(value):
    """
    Normalizuje nagłówek kolumny arkusza prowizji do porównań.
    """
    if value is None:
        return ''
    text = str(value).strip().upper()
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r'[^A-Z0-9]+', '', text)


def _find_header_col(row_map, predicate):
    """
    Zwraca indeks kolumny dla pierwszego nagłówka spełniającego warunek.
    """
    for key, col in row_map.items():
        if predicate(key):
            return col
    return None


def _pick_commission_rules_for_month(rules, month_start, month_end):
    """
    Wybiera najtrafniejsze rekordy prowizji dla miesiąca.

    Najpierw filtruje reguły nachodzące na miesiąc raportowy,
    a następnie bierze komplet rekordów z "najlepszego" okna:
    najpóźniejsze valid_from i (w tym zbiorze) najpóźniejsze valid_to.

    Dzięki temu jeżeli ryczałt i procent są zapisane w osobnych wierszach
    dla tego samego okresu, oba składniki są uwzględnione.
    """
    if not rules:
        return []

    matched = []
    for rule in rules:
        valid_from = rule.get('valid_from')
        valid_to = rule.get('valid_to')
        if valid_from is None or valid_to is None:
            continue
        if valid_from <= month_end and valid_to >= month_start:
            matched.append(rule)

    if not matched:
        return []

    best_valid_from = max((r.get('valid_from') or date.min) for r in matched)
    best_from_rules = [
        r for r in matched
        if (r.get('valid_from') or date.min) == best_valid_from
    ]
    best_valid_to = max((r.get('valid_to') or date.min) for r in best_from_rules)

    return [
        r for r in best_from_rules
        if (r.get('valid_to') or date.min) == best_valid_to
    ]


def _commission_amount_from_rule(rule, net_amount):
    """
    Wylicza prowizję dla jednej reguły i kwoty netto.
    """
    if not rule:
        return 0.0

    fixed_amount = float(rule.get('fixed_amount', 0.0) or 0.0)
    percent = float(rule.get('percent', 0.0) or 0.0)
    percent_amount = 0.0
    if percent > 0:
        percent_amount = float(net_amount or 0.0) * percent
    return max(fixed_amount, 0.0) + percent_amount


def _commission_amount_from_rules(rules, net_amount):
    """
    Sumuje prowizję ze wszystkich przekazanych reguł.
    """
    if not rules:
        return 0.0
    return sum(_commission_amount_from_rule(rule, net_amount) for rule in rules)

def build_monthly_commission_data_from_rules(revenue_data, commission_rules, month_str):
    """
    Buduje agregat prowizji per automat na podstawie reguł z pliku prowizji
    i obrotu per przewoźnik.

    Zwraca format zgodny z compare_commission_with_reference:
      {device_id: {'prowizja_zl': float, 'liczba_rekordow': int}}
    """
    month_start, month_end = _month_date_bounds(month_str)
    commission_data = {}
    missing_rules = 0
    missing_samples = []

    for device_id, rev_entry in (revenue_data or {}).items():
        if not isinstance(rev_entry, dict):
            continue
        by_carrier = rev_entry.get('by_carrier', {})
        if not isinstance(by_carrier, dict):
            continue

        device_id_int = int(device_id)
        total_device_commission = 0.0
        matched_records = 0

        for carrier_code, carrier_entry in by_carrier.items():
            gross_amount = 0.0
            net_amount = 0.0
            if isinstance(carrier_entry, dict):
                gross_amount = float(carrier_entry.get('obrot_brutto_zl', 0.0) or 0.0)
                net_amount = gross_amount / 1.08

            rules_for_key = commission_rules.get((device_id_int, carrier_code), [])
            picked_rules = _pick_commission_rules_for_month(rules_for_key, month_start, month_end)
            if picked_rules:
                matched_records += 1
            else:
                missing_rules += 1
                if len(missing_samples) < 8:
                    missing_samples.append((device_id_int, carrier_code))

            total_device_commission += _commission_amount_from_rules(picked_rules, net_amount)

        commission_data[device_id_int] = {
            'prowizja_zl': float(total_device_commission),
            'liczba_rekordow': int(matched_records),
        }

    if missing_rules:
        print(f"⚠ Brak dopasowanej reguły prowizji dla {missing_rules} par (automat, przewoźnik)")
        if missing_samples:
            print(f"  Przykłady braków: {missing_samples}")

    return commission_data


def load_commission_rules_and_locations_from_xlsx(month_str, commission_file=DEFAULT_PROWIZJE_FILE):
    """
    Wczytuje reguły prowizji i lokalizacje z pliku Prowizje_AB.xlsx.
    Zwraca:
      - commission_rules: {(device_id, carrier_code): [rule, ...]}
      - location_by_device: {device_id: location}
    """
    commission_path = Path(commission_file)
    if not commission_path.exists():
        print(f"⚠ Brak pliku prowizji: {commission_path}")
        return {}, {}

    try:
        from openpyxl import load_workbook
        wb = load_workbook(commission_path, data_only=True, read_only=True)
    except Exception as e:
        print(f"⚠ Nie udało się odczytać pliku prowizji: {e}")
        return {}, {}

    if not wb.sheetnames:
        return {}, {}

    ws = wb[wb.sheetnames[0]]
    header_row = None
    header_map = {}

    scan_to = min(ws.max_row, 40)
    for r in range(1, scan_to + 1):
        row_map = {}
        for c in range(1, ws.max_column + 1):
            token = _normalize_excel_commission_header(ws.cell(r, c).value)
            if not token:
                continue
            row_map[token] = c

        if not row_map:
            continue

        carrier_col = _find_header_col(row_map, lambda k: 'PRZEWOZNIK' in k)
        device_col = _find_header_col(row_map, lambda k: 'NUMERAUTOMATU' in k)
        valid_from_col = _find_header_col(
            row_map,
            lambda k: (
                'WAZNOSCOD' in k
                or 'WANOOD' in k
                or (k.startswith('WA') and k.endswith('OD'))
            ),
        )
        valid_to_col = _find_header_col(
            row_map,
            lambda k: (
                'WAZNOSCDO' in k
                or 'WANODO' in k
                or (k.startswith('WA') and k.endswith('DO'))
            ),
        )
        fixed_col = _find_header_col(
            row_map,
            lambda k: (
                'RYCZALT' in k
                or 'RYCZAT' in k
                or 'STALA' in k
                or 'KWOTOWA' in k
            ),
        )
        percent_col = _find_header_col(row_map, lambda k: 'PROCENTOWA' in k or 'PROCENT' in k)
        location_col = _find_header_col(row_map, lambda k: 'TVMLOCATION1' in k)

        if carrier_col and device_col and valid_from_col and valid_to_col and (fixed_col or percent_col):
            header_row = r
            header_map = {
                'carrier_col': carrier_col,
                'device_col': device_col,
                'valid_from_col': valid_from_col,
                'valid_to_col': valid_to_col,
                'fixed_col': fixed_col,
                'percent_col': percent_col,
                'location_col': location_col,
            }
            break

    if header_row is None:
        print("⚠ Nie znaleziono wymaganych kolumn w pliku prowizji")
        return {}, {}

    commission_rules = {}
    location_by_device = {}
    month_start, month_end = _month_date_bounds(month_str)
    invalid_date_ranges = 0
    skipped_missing_required = 0
    skipped_zero_after_parse = 0
    fixed_only_rules = 0
    percent_only_rules = 0
    mixed_rules = 0
    parse_to_zero_fixed = 0
    parse_to_zero_percent = 0
    parse_to_zero_samples = []

    for r in range(header_row + 1, ws.max_row + 1):
        raw_carrier = ws.cell(r, header_map['carrier_col']).value
        raw_device = ws.cell(r, header_map['device_col']).value
        raw_from = ws.cell(r, header_map['valid_from_col']).value
        raw_to = ws.cell(r, header_map['valid_to_col']).value

        if raw_carrier is None and raw_device is None and raw_from is None and raw_to is None:
            continue

        carrier_code = _normalize_carrier_code(raw_carrier)
        device_id = _parse_device_id(raw_device)
        valid_from = _as_date(raw_from)
        valid_to = _as_date(raw_to)

        if not carrier_code or device_id is None or valid_from is None or valid_to is None:
            skipped_missing_required += 1
            continue

        if valid_from > valid_to:
            invalid_date_ranges += 1
            continue

        fixed_amount = 0.0
        percent = 0.0
        raw_fixed = None
        raw_percent = None
        if header_map.get('fixed_col'):
            raw_fixed = ws.cell(r, header_map['fixed_col']).value
            fixed_amount = _as_float(raw_fixed)
        if header_map.get('percent_col'):
            raw_percent = ws.cell(r, header_map['percent_col']).value
            percent = _as_float(raw_percent, as_percent=True)

        if _has_nonempty_value(raw_fixed) and fixed_amount == 0.0 and re.search(r'[1-9]', str(raw_fixed)):
            parse_to_zero_fixed += 1
            if len(parse_to_zero_samples) < 8:
                parse_to_zero_samples.append((r, 'fixed', str(raw_fixed)))

        if _has_nonempty_value(raw_percent) and percent == 0.0 and re.search(r'[1-9]', str(raw_percent)):
            parse_to_zero_percent += 1
            if len(parse_to_zero_samples) < 8:
                parse_to_zero_samples.append((r, 'percent', str(raw_percent)))

        if fixed_amount <= 0.0 and percent <= 0.0:
            skipped_zero_after_parse += 1
            continue

        if fixed_amount > 0.0 and percent > 0.0:
            mixed_rules += 1
        elif fixed_amount > 0.0:
            fixed_only_rules += 1
        else:
            percent_only_rules += 1

        key = (int(device_id), carrier_code)
        commission_rules.setdefault(key, []).append({
            'valid_from': valid_from,
            'valid_to': valid_to,
            'fixed_amount': fixed_amount,
            'percent': percent,
        })

        raw_location = ws.cell(r, header_map['location_col']).value if header_map.get('location_col') else None
        normalized_location = _normalize_tvm_location_text(raw_location)
        if (
            normalized_location
            and int(device_id) not in location_by_device
            and valid_from <= month_end
            and valid_to >= month_start
        ):
            location_by_device[int(device_id)] = normalized_location

    if invalid_date_ranges:
        print(
            f"⚠ Pominięto {invalid_date_ranges} wierszy prowizji z nieprawidłowym zakresem dat (od > do)"
        )
    if skipped_missing_required:
        print(
            f"⚠ Pominięto {skipped_missing_required} wierszy prowizji z brakiem wymaganych danych "
            "(przewoźnik/automat/daty)"
        )
    if skipped_zero_after_parse:
        print(
            f"⚠ Pominięto {skipped_zero_after_parse} wierszy prowizji, bo ryczałt i procent wyszły jako 0"
        )

    parse_to_zero_total = parse_to_zero_fixed + parse_to_zero_percent
    if parse_to_zero_total:
        print(
            "⚠ Podejrzane parse-to-zero w prowizjach: "
            f"fixed={parse_to_zero_fixed}, percent={parse_to_zero_percent}"
        )
        if parse_to_zero_samples:
            print(f"  Przykłady parse-to-zero: {parse_to_zero_samples}")

    print(
        "✓ Typy reguł prowizji: "
        f"ryczałt={fixed_only_rules}, "
        f"procent={percent_only_rules}, "
        f"mieszane={mixed_rules}"
    )
    loaded_carriers = sorted({carrier for _, carrier in commission_rules.keys()})
    if loaded_carriers:
        print(f"✓ Przewoźnicy w regułach prowizji: {loaded_carriers}")

    return commission_rules, location_by_device


def _build_month_header_targets(month_str):
    """
    Buduje zestaw tokenów nagłówka miesiąca dla wyszukiwania kolumn w Excel.
    """
    dt = datetime.strptime(month_str, '%Y-%m')
    month_str_slash = dt.strftime('%m/%Y')
    month_str_dot = dt.strftime('%m.%Y')
    month_str_dash = dt.strftime('%Y-%m')
    month_str_compact = dt.strftime('%m%Y')

    tokens = {
        _normalize_header_text(month_str_slash),
        _normalize_header_text(month_str_dot),
        _normalize_header_text(month_str_dash),
        _normalize_header_text(month_str_compact),
        _normalize_header_text(f"Amortyzacja {month_str_slash}"),
        _normalize_header_text(f"Amortyzacja {month_str_dot}"),
        _normalize_header_text(f"Amortyzacja {month_str_dash}"),
        _normalize_header_text(f"Amortyzacja {month_str_compact}"),
    }
    return {token for token in tokens if token}


def _load_bb_amortyzacja_sources(
    month_str,
    amortyzacja_file=DEFAULT_AMORTYZACJA_FILE,
    amortyzacja_sheet=DEFAULT_AMORTYZACJA_SHEET,
):
    """
    Wczytuje źródła amortyzacji BB z arkusza bb8:
    - mapę po nazwie lokalizacji (kolumna Nazwa)
    - mapę po numerze inwentarzowym (kolumna Nr.inw)
    Kolumna kwoty jest wykrywana dynamicznie po miesiącu raportu.
    """
    result = {
        'by_location': {},
        'by_nr_inw': {},
        'stats': {
            'rows_read': 0,
            'duplicates_location': 0,
            'duplicates_nr_inw': 0,
        },
    }
    if amortyzacja_file is None:
        return result

    amort_path = Path(amortyzacja_file)
    if not amort_path.exists():
        print(f"⚠ Brak pliku amortyzacji: {amort_path}")
        return result

    try:
        from openpyxl import load_workbook
        wb = load_workbook(amort_path, data_only=True, read_only=True)
    except Exception as e:
        print(f"⚠ Nie udało się odczytać pliku amortyzacji: {e}")
        return result

    sheet_name = str(amortyzacja_sheet or DEFAULT_AMORTYZACJA_SHEET)
    if sheet_name not in wb.sheetnames:
        print(f"⚠ Brak arkusza {sheet_name} w pliku amortyzacji: {amort_path.name}")
        return result

    ws = wb[sheet_name]
    month_tokens = _build_month_header_targets(month_str)

    header_row = None
    nazwa_col = None
    nr_inw_col = None
    amount_col = None

    scan_to = min(ws.max_row, 60)
    for r in range(1, scan_to + 1):
        row_nazwa_col = None
        row_nr_inw_col = None
        row_amount_col = None

        for c in range(1, ws.max_column + 1):
            normalized = _normalize_header_text(ws.cell(r, c).value)
            if not normalized:
                continue

            if normalized in {'NAZWA'}:
                row_nazwa_col = c
            if normalized in {'NRINW', 'NRINWENTARZOWY'}:
                row_nr_inw_col = c
            if (
                normalized in month_tokens
                or any(token and token in normalized for token in month_tokens)
            ):
                row_amount_col = c

        if row_nazwa_col and row_amount_col:
            header_row = r
            nazwa_col = row_nazwa_col
            nr_inw_col = row_nr_inw_col
            amount_col = row_amount_col
            break

    if header_row is None:
        print(
            "⚠ Nie znaleziono kolumn amortyzacji BB (Nazwa + miesiąc) "
            f"w arkuszu {sheet_name}"
        )
        return result

    if nazwa_col is None or amount_col is None:
        print(f"⚠ Niekompletna konfiguracja kolumn amortyzacji BB w arkuszu {sheet_name}")
        return result

    nazwa_col = int(nazwa_col)
    amount_col = int(amount_col)
    if nr_inw_col is not None:
        nr_inw_col = int(nr_inw_col)

    for r in range(header_row + 1, ws.max_row + 1):
        raw_nazwa = ws.cell(r, nazwa_col).value
        raw_amount = ws.cell(r, amount_col).value
        raw_nr_inw = ws.cell(r, nr_inw_col).value if nr_inw_col else None

        if raw_nazwa is None and raw_amount is None and raw_nr_inw is None:
            continue

        result['stats']['rows_read'] += 1
        amount = _as_float(raw_amount)

        normalized_name = _normalize_text_for_match(raw_nazwa)
        if normalized_name:
            if normalized_name in result['by_location']:
                result['stats']['duplicates_location'] += 1
            result['by_location'][normalized_name] = amount

        parsed_nr_inw = _parse_device_id(raw_nr_inw)
        if parsed_nr_inw is not None:
            if parsed_nr_inw in result['by_nr_inw']:
                result['stats']['duplicates_nr_inw'] += 1
            result['by_nr_inw'][parsed_nr_inw] = amount

    return result


def _load_rent_costs_by_device(month_str, rent_file=DEFAULT_RENT_FILE, rent_sheet=None):
    """
    Wczytuje czynsz i prąd per automat z pliku najmu.
    Oczekuje kolumn: Nr.automatu, Czynsz MM/YYYY, Energia MM/YYYY.
    """
    if rent_file is None:
        return {}

    rent_path = Path(rent_file)
    if not rent_path.exists():
        print(f"⚠ Brak pliku najmu: {rent_path}")
        return {}

    try:
        from openpyxl import load_workbook
        wb = load_workbook(rent_path, data_only=True, read_only=True)
    except Exception as e:
        print(f"⚠ Nie udało się odczytać pliku najmu: {e}")
        return {}

    year_sheet = month_str.split('-')[0]
    sheet_name = str(rent_sheet or year_sheet)
    if sheet_name not in wb.sheetnames:
        print(f"⚠ Brak arkusza {sheet_name} w pliku najmu: {rent_path.name}")
        return {}

    try:
        ws = wb[sheet_name]
        month_label_slash, month_label_compact = _build_month_year_labels(month_str)

        header_row = None
        device_col = None
        czynsz_col = None
        prad_col = None

        czynsz_targets = {
            _normalize_header_text(f"Czynsz {month_label_slash}"),
            _normalize_header_text(f"Czynsz{month_label_slash}"),
            _normalize_header_text(f"Czynsz {month_label_compact}"),
            _normalize_header_text(f"Czynsz{month_label_compact}"),
        }
        prad_targets = {
            _normalize_header_text(f"Energia {month_label_slash}"),
            _normalize_header_text(f"Energia{month_label_slash}"),
            _normalize_header_text(f"Energia {month_label_compact}"),
            _normalize_header_text(f"Energia{month_label_compact}"),
            _normalize_header_text(f"Prad {month_label_slash}"),
            _normalize_header_text(f"Prad{month_label_slash}"),
            _normalize_header_text(f"Prad {month_label_compact}"),
            _normalize_header_text(f"Prad{month_label_compact}"),
        }

        scan_to = min(ws.max_row, 60)
        for r, row_values in enumerate(
            ws.iter_rows(min_row=1, max_row=scan_to, values_only=True),
            start=1,
        ):
            row_device_col = None
            row_czynsz_col = None
            row_prad_col = None

            for c, raw_header in enumerate(row_values, start=1):
                normalized_header = _normalize_header_text(raw_header)
                if not normalized_header:
                    continue

                if normalized_header == 'NRAUTOMATU':
                    row_device_col = c
                if normalized_header in czynsz_targets:
                    row_czynsz_col = c
                if normalized_header in prad_targets:
                    row_prad_col = c

            if row_device_col and row_czynsz_col and row_prad_col:
                header_row = r
                device_col = row_device_col
                czynsz_col = row_czynsz_col
                prad_col = row_prad_col
                break

        if header_row is None:
            print(
                "⚠ Nie znaleziono wymaganych kolumn najmu: "
                f"Nr.automatu, Czynsz {month_label_slash}, Energia {month_label_slash}"
            )
            return {}

        if device_col is None or czynsz_col is None or prad_col is None:
            print("⚠ Niekompletna konfiguracja kolumn najmu")
            return {}

        min_col = min(device_col, czynsz_col, prad_col)
        max_col = max(device_col, czynsz_col, prad_col)
        device_idx = device_col - min_col
        czynsz_idx = czynsz_col - min_col
        prad_idx = prad_col - min_col

        rent_map = {}
        skipped_rows = 0
        empty_streak = 0

        for row_values in ws.iter_rows(
            min_row=header_row + 1,
            max_row=ws.max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=True,
        ):
            raw_device = row_values[device_idx]
            raw_czynsz = row_values[czynsz_idx]
            raw_prad = row_values[prad_idx]

            if raw_device is None and raw_czynsz is None and raw_prad is None:
                empty_streak += 1
                if empty_streak >= 200:
                    break
                continue

            empty_streak = 0

            device_id = _parse_device_id(raw_device)
            if device_id is None:
                skipped_rows += 1
                continue

            rent_map[device_id] = {
                'czynsz': _as_float(raw_czynsz),
                'prad': _as_float(raw_prad),
            }

        if skipped_rows:
            print(f"⚠ Pominięto {skipped_rows} wierszy najmu z nieprawidłowym Nr.automatu")

        return rent_map
    finally:
        wb.close()


def _get_monthly_costs_per_device(
    month_str,
    device_ids,
    costs_file=DEFAULT_COSTS_FILE,
    rent_file=DEFAULT_RENT_FILE,
    rent_sheet=None,
):
    """
    Wczytuje koszty z pliku Excel i rozdziela je na automaty.
    Dla kosztów globalnych stosuje równy podział na wszystkie automaty AB.
    """
    normalized_ids = sorted({int(device_id) for device_id in device_ids})
    if not normalized_ids:
        return {}

    default_entry = {
        **{key: 0.0 for key in TVM_COST_KEYS},
        **{key: 0.0 for key in OTHER_COST_KEYS},
    }

    if costs_file is None:
        return {device_id: dict(default_entry) for device_id in normalized_ids}

    costs_path = Path(costs_file)
    if not costs_path.exists():
        print(f"⚠ Brak pliku kosztów: {costs_path}")
        return {device_id: dict(default_entry) for device_id in normalized_ids}

    try:
        from openpyxl import load_workbook
        wb = load_workbook(costs_path, data_only=True, read_only=True)
    except Exception as e:
        print(f"⚠ Nie udało się odczytać pliku kosztów: {e}")
        return {device_id: dict(default_entry) for device_id in normalized_ids}

    year_sheet = month_str.split('-')[0]
    if year_sheet not in wb.sheetnames:
        print(f"⚠ Brak arkusza {year_sheet} w pliku kosztów: {costs_path.name}")
        return {device_id: dict(default_entry) for device_id in normalized_ids}

    ws = wb[year_sheet]
    global_costs = dict(default_entry)
    rent_by_device = _load_rent_costs_by_device(
        month_str,
        rent_file=rent_file,
        rent_sheet=rent_sheet,
    )

    # Koszty TVM (sekcje w kolumnie B)
    global_costs['elavon'] = _extract_monthly_value_in_section(ws, 'ELAVON', month_str)
    global_costs['poczta_polska'] = _extract_monthly_value_in_section(ws, 'POCZTA POLSKA', month_str)
    global_costs['amortyzacja'] = _extract_monthly_value_in_section(ws, 'AMORTYZACJA', month_str)
    global_costs['papier'] = _extract_monthly_value_in_section(ws, 'PAPIER', month_str)
    global_costs['transmisja_danych'] = _extract_monthly_value_in_section(ws, 'TELEFONY/INTERNET', month_str)
    global_costs['utrzymanie_oprogramowania'] = _extract_monthly_value_in_section(ws, 'NORDPLUS', month_str)
    global_costs['ubezpieczenie'] = _extract_monthly_value_in_section(ws, 'UBEZPIECZENIE AUTOMATÓW', month_str)

    serwis_label = 'SERWIS TERENOWY (AB, bez rPOS)'
    serwis_row = _find_section_row(ws, serwis_label)
    serwis_end_row = _get_section_end_row(ws, serwis_row)
    serwis_glowny = _extract_monthly_value_in_section(ws, serwis_label, month_str)
    serwis_szyszkowska = _extract_monthly_value_from_columns(
        ws,
        month_str,
        month_col=9,
        value_col=10,
        start_row=(serwis_row + 1) if serwis_row else 1,
        end_row=serwis_end_row,
    )
    serwis_zdankiewicz = _extract_monthly_value_from_columns(
        ws,
        month_str,
        month_col=14,
        value_col=15,
        start_row=(serwis_row + 1) if serwis_row else 1,
        end_row=serwis_end_row,
    )
    global_costs['serwis'] = serwis_glowny + serwis_szyszkowska
    global_costs['zdankiewicz'] = serwis_zdankiewicz

    # Czynsz i prąd są pobierane per automat z pliku najmu.
    global_costs['czynsz'] = 0.0
    global_costs['prad'] = 0.0
    global_costs['it_card'] = 0.0

    # Other project costs
    global_costs['non_tvm'] = _extract_monthly_value_from_header_table(
        ws,
        month_str,
        header_label='Non TVMs',
        month_col=9,
        value_col=10,
    )
    global_costs['project_variable_costs'] = _extract_monthly_value_from_header_table(
        ws,
        month_str,
        header_label='Project',
        month_col=9,
        value_col=11,
    )
    # OH będzie liczony per-device poniżej

    device_count = len(normalized_ids)
    costs_by_device = {}
    for device_id in normalized_ids:
        per_device = {}
        for key, value in global_costs.items():
            if key == 'oh':
                continue
            per_device[key] = float(value or 0.0) / device_count

        rent_entry = rent_by_device.get(device_id, {})
        per_device['czynsz'] = float(rent_entry.get('czynsz', 0.0) or 0.0)
        per_device['prad'] = float(rent_entry.get('prad', 0.0) or 0.0)
        
        # Calculate OH per device: (NON_TVM + Project_variable_costs + TVM_sum) * 0.2
        tvm_sum = sum(float(per_device.get(key, 0.0) or 0.0) for key in TVM_COST_KEYS)
        per_device['oh'] = (
            float(per_device.get('non_tvm', 0.0) or 0.0) +
            float(per_device.get('project_variable_costs', 0.0) or 0.0) +
            tvm_sum
        ) * 0.2
        
        costs_by_device[device_id] = per_device

    return costs_by_device


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
    Normalizuje kod przewoźnika do postaci technicznej używanej w kluczach.
    Usuwa diakrytyki, np. KMŁ -> KML.
    """
    normalized = str(carrier or '').strip().upper()
    normalized = unicodedata.normalize('NFKD', normalized)
    normalized = ''.join(ch for ch in normalized if not unicodedata.combining(ch))
    return normalized


def _build_amount_expr_sql(amount_col_name):
    """
    Buduje wyrażenie SQL kwoty obrotu.
    Dla pól finansowych monitora przechowywanych w groszach zwraca wartość w złotych.
    """
    amount_col_sql = sql.Identifier(amount_col_name)
    if amount_col_name.lower() in {'fin_nalezn', 'fin_ptu_kwota'}:
        return sql.SQL("(COALESCE({amount_col}, 0)::numeric / 100.0)").format(
            amount_col=amount_col_sql
        )
    return sql.SQL("COALESCE({amount_col}, 0)").format(amount_col=amount_col_sql)


def _display_carrier_label(carrier_code):
    """
    Zwraca etykietę przewoźnika do nagłówka raportu.
    """
    normalized = _normalize_carrier_code(carrier_code)
    if normalized == 'KML':
        return 'KMŁ'
    return normalized


def _resolve_transactions_table_for_carrier(carrier_code, default_table):
    """
    Zwraca nazwę tabeli transakcji dla przewoźnika.
    Pozwala nadpisać domyślną tabelę dla wybranych schematów.
    """
    normalized = _normalize_carrier_code(carrier_code)
    return TRANSACTIONS_TABLE_OVERRIDES.get(normalized, default_table)


def _resolve_transactions_device_col_for_carrier(carrier_code, default_device_col):
    """
    Zwraca kolumnę urządzenia dla przewoźnika (string lub lista fallbacków).
    """
    normalized = _normalize_carrier_code(carrier_code)
    return TRANSACTIONS_DEVICE_COL_OVERRIDES.get(normalized, default_device_col)


def _resolve_transactions_amount_col_for_carrier(carrier_code, default_amount_col):
    """
    Zwraca kolumnę kwoty dla przewoźnika (string lub lista fallbacków).
    """
    normalized = _normalize_carrier_code(carrier_code)
    return TRANSACTIONS_AMOUNT_COL_OVERRIDES.get(normalized, default_amount_col)


def _resolve_transactions_payment_method_col_for_carrier(carrier_code, default_payment_method_col):
    """
    Zwraca kolumnę metody płatności dla przewoźnika (string lub lista fallbacków).
    """
    normalized = _normalize_carrier_code(carrier_code)
    return TRANSACTIONS_PAYMENT_METHOD_COL_OVERRIDES.get(normalized, default_payment_method_col)


def _preview_column_name(column_ref):
    """
    Zwraca nazwę kolumny do podglądu SQL (dla list fallbacków używa pierwszej pozycji).
    """
    if isinstance(column_ref, (list, tuple)):
        return column_ref[0] if column_ref else ''
    return column_ref


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
    payment_method_col=None,
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

    if isinstance(amount_col, (list, tuple)):
        picked_amount_col = _pick_first_matching_column(columns, list(amount_col))
    else:
        picked_amount_col = lowered_map.get(str(amount_col).lower())

    if isinstance(date_col, (list, tuple)):
        picked_date_col = _pick_first_matching_column(columns, list(date_col))
    else:
        picked_date_col = lowered_map.get(str(date_col).lower())

    if picked_amount_col is None or picked_date_col is None:
        return None

    if device_col:
        if isinstance(device_col, (list, tuple)):
            picked_device_col = _pick_first_matching_column(columns, list(device_col))
        else:
            picked_device_col = lowered_map.get(str(device_col).lower())
    else:
        picked_device_col = _pick_first_matching_column(columns, TRANSACTIONS_DEVICE_COL_CANDIDATES)

    if picked_device_col is None:
        return None

    if payment_method_col:
        if isinstance(payment_method_col, (list, tuple)):
            picked_payment_method_col = _pick_first_matching_column(columns, list(payment_method_col))
        else:
            picked_payment_method_col = lowered_map.get(str(payment_method_col).lower())
    else:
        picked_payment_method_col = _pick_first_matching_column(
            columns,
            TRANSACTIONS_PAYMENT_METHOD_COL_CANDIDATES,
        )

    if picked_payment_method_col is None:
        return None

    return {
        'schema': schema_name,
        'table': table_name,
        'device_col': picked_device_col,
        'amount_col': picked_amount_col,
        'date_col': picked_date_col,
        'payment_method_col': picked_payment_method_col,
    }


def build_transactions_debug_query(
    schema_name,
    month_str,
    table_name=TRANSACTIONS_SOURCE_CONFIG['table'],
    amount_col=TRANSACTIONS_SOURCE_CONFIG['amount_col'],
    date_col=TRANSACTIONS_SOURCE_CONFIG['date_col'],
    device_col=TRANSACTIONS_SOURCE_CONFIG['device_col'],
    payment_method_col=TRANSACTIONS_SOURCE_CONFIG['payment_method_col'],
):
    """
    Buduje gotowe zapytanie SQL diagnostyczne dla jednego schematu przewoźnika.
    """
    amount_col = _preview_column_name(amount_col)
    date_col = _preview_column_name(date_col)
    device_col = _preview_column_name(device_col)
    payment_method_col = _preview_column_name(payment_method_col)
    start_date, end_date = get_month_range_closed(month_str)
    amount_expr = f"COALESCE({amount_col}, 0)"
    if amount_col.lower() in {'fin_nalezn', 'fin_ptu_kwota'}:
        amount_expr = f"(COALESCE({amount_col}, 0)::numeric / 100.0)"
    return f"""
SELECT
    {device_col} AS device_id,
    {payment_method_col} AS fin_spos_opl,
    SUM({amount_expr}) AS obrot_metoda_zl,
    COUNT(*) AS liczba_transakcji
FROM {schema_name}.{table_name}
WHERE {device_col}::text ~ '^[0-9]+$'
  AND (({device_col}::bigint BETWEEN 1101 AND 1141) OR ({device_col}::bigint BETWEEN 1201 AND 1299))
  AND {date_col} >= TIMESTAMP '{start_date.strftime('%Y-%m-%d %H:%M:%S')}'
  AND {date_col} <= TIMESTAMP '{end_date.strftime('%Y-%m-%d %H:%M:%S')}'
    AND {payment_method_col}::text IN ('1', '2', '6')
GROUP BY {device_col}, {payment_method_col}
ORDER BY device_id, {payment_method_col};
""".strip()


def get_monthly_revenue_by_carrier(
    conn,
    month_str,
    carriers=None,
    table_name=TRANSACTIONS_SOURCE_CONFIG['table'],
    amount_col=TRANSACTIONS_SOURCE_CONFIG['amount_col'],
    date_col=TRANSACTIONS_SOURCE_CONFIG['date_col'],
    device_col=TRANSACTIONS_SOURCE_CONFIG['device_col'],
    payment_method_col=TRANSACTIONS_SOURCE_CONFIG['payment_method_col'],
    use_device_range_filter=True,
):
    """
    Pobiera obrót brutto per automat i per przewoźnik z tabeli <carrier_schema>.transactions.
    Zwraca: dict {
        device_id: {
            'by_carrier': {
                carrier: {
                    'obrot_brutto_zl': float,
                    'liczba_transakcji': int,
                    'by_payment_method': {
                        'gotowka': float,
                        'karta': float,
                        'blik': float,
                    }
                }
            }
        }
    }
    """
    start_date, end_date = get_month_range_closed(month_str)
    carriers = carriers or CARRIERS
    expected_carriers = sorted({_normalize_carrier_code(carrier) for carrier in carriers})

    revenue_data = {}

    for carrier in carriers:
        carrier_code = _normalize_carrier_code(carrier)
        carrier_table_name = _resolve_transactions_table_for_carrier(carrier_code, table_name)
        carrier_device_col = _resolve_transactions_device_col_for_carrier(carrier_code, device_col)
        carrier_amount_col = _resolve_transactions_amount_col_for_carrier(carrier_code, amount_col)
        carrier_payment_method_col = _resolve_transactions_payment_method_col_for_carrier(
            carrier_code,
            payment_method_col,
        )
        source = _resolve_transactions_source(
            conn,
            schema_name=carrier_code,
            table_name=carrier_table_name,
            amount_col=carrier_amount_col,
            date_col=date_col,
            device_col=carrier_device_col,
            payment_method_col=carrier_payment_method_col,
        )
        if source is None:
            print(
                f"⚠ Pomijam przewoźnika {carrier} (schemat {carrier_code}): "
                f"brak źródła {carrier_code}.{carrier_table_name} lub wymaganych kolumn "
                f"({carrier_device_col}, {date_col}, {carrier_amount_col}, {carrier_payment_method_col})"
            )
            continue

        params = [start_date, end_date]
        device_filter_sql = sql.SQL('')
        if use_device_range_filter:
            device_expr = sql.SQL("{device_col}::bigint").format(
                device_col=sql.Identifier(source['device_col'])
            )
            device_filter_sql = sql.SQL(" AND {device_filter}").format(
                device_filter=_build_device_filter_sql(device_expr)
            )
        numeric_device_sql = sql.SQL(" AND {device_col}::text ~ '^[0-9]+$'").format(
            device_col=sql.Identifier(source['device_col'])
        )

        query = sql.SQL(
            """
            SELECT
                {device_col} AS device_id,
                {payment_method_col} AS fin_spos_opl,
                SUM({amount_expr}) AS obrot_metoda_zl,
                COUNT(*) AS liczba_transakcji
            FROM {schema}.{table}
            WHERE {date_col} >= %s
                            AND {date_col} <= %s
                            {numeric_device_filter}
              AND {payment_method_col}::text IN ('1', '2', '6')
              {device_filter}
            GROUP BY {device_col}, {payment_method_col}
            ORDER BY device_id, {payment_method_col}
            """
        ).format(
            schema=sql.Identifier(source['schema']),
            table=sql.Identifier(source['table']),
            device_col=sql.Identifier(source['device_col']),
            amount_expr=_build_amount_expr_sql(source['amount_col']),
            date_col=sql.Identifier(source['date_col']),
            payment_method_col=sql.Identifier(source['payment_method_col']),
            numeric_device_filter=numeric_device_sql,
            device_filter=device_filter_sql,
         )

        cursor = conn.cursor()
        try:
            cursor.execute(query, tuple(params))
            rows = cursor.fetchall()
            print(
                f"✓ Obrót {carrier_code}: {source['schema']}.{source['table']} "
                f"({source['device_col']}, {source['amount_col']}, {source['date_col']}, {source['payment_method_col']}) -> {len(rows)} rekordów"
            )
        except psycopg2.Error as e:
            print(f"⚠ Błąd pobierania obrotu dla {carrier_code}: {e}")
            rows = []
        finally:
            cursor.close()

        for row in rows:
            device_id, payment_method_value, obrot, liczba = row
            try:
                device_id_int = int(device_id)
            except (TypeError, ValueError):
                continue

            method_key = PAYMENT_METHOD_CODE_TO_KEY.get(str(payment_method_value).strip())
            if method_key is None:
                continue

            revenue_data.setdefault(device_id_int, {'by_carrier': {}})
            carrier_entry = revenue_data[device_id_int]['by_carrier'].setdefault(
                carrier_code,
                {
                    'obrot_brutto_zl': 0.0,
                    'liczba_transakcji': 0,
                    'by_payment_method': {key: 0.0 for key in PAYMENT_METHOD_KEYS},
                },
            )

            method_amount = float(obrot or 0.0)
            method_count = int(liczba or 0)
            carrier_entry['by_payment_method'][method_key] += method_amount
            carrier_entry['obrot_brutto_zl'] += method_amount
            carrier_entry['liczba_transakcji'] += method_count

        carrier_total = sum(
            rev_entry.get('obrot_brutto_zl', 0.0)
            for rev_device in revenue_data.values()
            for code, rev_entry in rev_device.get('by_carrier', {}).items()
            if code == carrier_code
        )
        print(f"  Suma {carrier_code}: {carrier_total:,.2f} zł")

    observed_carriers = sorted({
        code
        for rev_device in revenue_data.values()
        for code in rev_device.get('by_carrier', {}).keys()
    })
    print(f"✓ Przewoźnicy w obrocie: {observed_carriers}")
    missing_expected = sorted(set(expected_carriers) - set(observed_carriers))
    if missing_expected:
        print(f"⚠ Brak danych obrotu dla przewoźników: {missing_expected}")

    return revenue_data


# === EKSPORT DO EXCEL (KOLUMNY 1-2, RESZTA PLACEHOLDER) ===

def export_to_excel_PL(
    dictionary_comparison,
    revenue_data,
    month_str,
    filename,
    commission_rules=None,
    carriers=None,
    location_by_device=None,
    automat_type_by_device=None,
    costs_file=DEFAULT_COSTS_FILE,
    rent_file=DEFAULT_RENT_FILE,
    rent_sheet=None,
    amortyzacja_file=DEFAULT_AMORTYZACJA_FILE,
    amortyzacja_sheet=DEFAULT_AMORTYZACJA_SHEET,
):
    """
    Eksportuje raport P&L do Excela.
    Układ danych: jeden wiersz per automat, kolumny brutto i metody płatności per przewoźnik.
    """
    if commission_rules is None:
        commission_rules = {}

    carrier_order = []
    for carrier in (carriers or CARRIERS):
        code = _normalize_carrier_code(carrier)
        if code not in carrier_order:
            carrier_order.append(code)

    for rev_entry in revenue_data.values():
        by_carrier = rev_entry.get('by_carrier') if isinstance(rev_entry, dict) else None
        if isinstance(by_carrier, dict):
            for code in by_carrier.keys():
                if code not in carrier_order:
                    carrier_order.append(code)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    filepath = OUTPUT_DIR / filename
    
    wb = Workbook()
    ws = wb.active
    ws.title = f"P&L {month_str}"
    
    # === NAGŁÓWKI ===
    headers = ['Nr aut.', 'Lokalizacja automatu', 'Typ automatu']
    for code in carrier_order:
        label = _display_carrier_label(code)
        headers.append(f"Brutto {label}")
        headers.append(f"Prowizja {label}")
        headers.append(f"Netto {label}")
        headers.append(f"{PAYMENT_METHOD_LABELS['gotowka']} {label}")
        headers.append(f"{PAYMENT_METHOD_LABELS['karta']} {label}")
        headers.append(f"{PAYMENT_METHOD_LABELS['blik']} {label}")
    headers.append('Brutto Suma')
    headers.append('Transakcje bezgotówkowe Suma')
    headers.append('Prowizja Suma')
    headers.append('Dodatkowe zyski')
    headers.append('Netto Suma')

    for key in TVM_COST_KEYS:
        headers.append(TVM_COST_LABELS[key])
    headers.append('Koszty TVM Suma')

    for key in OTHER_COST_KEYS:
        headers.append(OTHER_COST_LABELS[key])
    headers.append('TOTAL other project costs')
    headers.append('SUMA KOSZTY')

    headers.append('Karta Suma')
    headers.append('BLIK Suma')
    headers.append('Gotówka Suma')
    headers.append('Transakcje bezgotówkowe')
    headers.append('Wynik (Prowizja - Koszty)')
    headers.append('UWAGI')
    result_col_idx = headers.index('Wynik (Prowizja - Koszty)') + 1
    uwagi_col_idx = len(headers)
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # === DANE ===
    row_num = 2
    location_by_device = location_by_device or {}
    automat_type_by_device = automat_type_by_device or {}
    # Upewnij się, że wszystkie klucze są integerami
    all_device_ids = sorted(set(
        [int(k) for k in dictionary_comparison.keys()] + 
        [int(k) for k in revenue_data.keys()]
    ))
    costs_by_device = _get_monthly_costs_per_device(
        month_str,
        all_device_ids,
        costs_file=costs_file,
        rent_file=rent_file,
        rent_sheet=rent_sheet,
    )
    bb_amort_sources = _load_bb_amortyzacja_sources(
        month_str,
        amortyzacja_file=amortyzacja_file,
        amortyzacja_sheet=amortyzacja_sheet,
    )
    bb_type_count = 0
    bb_matched_by_location = 0
    bb_matched_by_nr = 0
    bb_missing = 0
    
    for device_id in all_device_ids:
        # 1-INFO (najpierw mapa lokalizacji zebrana po przewoźnikach)
        lokalizacja = location_by_device.get(device_id, '')
        if not lokalizacja and device_id in dictionary_comparison:
            info = dictionary_comparison[device_id]
            data = info['data']
            lokalizacja = extract_city_name(data.get('description', ''))
        
        rev_by_carrier = {}
        if device_id in revenue_data:
            rev_entry = revenue_data[device_id]
            if isinstance(rev_entry, dict) and isinstance(rev_entry.get('by_carrier'), dict):
                rev_by_carrier = rev_entry['by_carrier']
            elif isinstance(rev_entry, dict):
                if carrier_order:
                    rev_by_carrier = {carrier_order[0]: rev_entry}

        ws.cell(row=row_num, column=1).value = device_id
        ws.cell(row=row_num, column=2).value = lokalizacja
        automat_type = str(automat_type_by_device.get(device_id, '') or '').strip().upper()
        ws.cell(row=row_num, column=3).value = automat_type

        row_total = 0.0
        row_commission_total = 0.0
        row_netto_total = 0.0
        cash_total = 0.0
        card_total = 0.0
        blik_total = 0.0
        carrier_cashless_total = 0.0
        carrier_cashless_by_code = {}
        month_start, month_end = _month_date_bounds(month_str)
        col_idx = 4
        for carrier_code in carrier_order:
            rev = rev_by_carrier.get(carrier_code)
            amount = 0.0
            netto_amount = 0.0
            by_payment_method = {key: 0.0 for key in PAYMENT_METHOD_KEYS}
            if isinstance(rev, dict):
                amount = float(rev.get('obrot_brutto_zl', 0.0) or 0.0)
                method_values = rev.get('by_payment_method')
                if isinstance(method_values, dict):
                    for method_key in PAYMENT_METHOD_KEYS:
                        by_payment_method[method_key] = float(method_values.get(method_key, 0.0) or 0.0)
            netto_amount = amount / 1.08

            cash_total += by_payment_method['gotowka']
            card_total += by_payment_method['karta']
            blik_total += by_payment_method['blik']
            cashless_for_carrier = by_payment_method['karta'] + by_payment_method['blik']
            carrier_cashless_total += cashless_for_carrier
            carrier_cashless_by_code[carrier_code] = cashless_for_carrier

            rules_for_key = commission_rules.get((int(device_id), carrier_code), [])
            picked_rules = _pick_commission_rules_for_month(rules_for_key, month_start, month_end)
            commission_amount = _commission_amount_from_rules(picked_rules, netto_amount)

            ws.cell(row=row_num, column=col_idx).value = amount
            ws.cell(row=row_num, column=col_idx).number_format = '#,##0.00'
            ws.cell(row=row_num, column=col_idx + 1).value = commission_amount
            ws.cell(row=row_num, column=col_idx + 1).number_format = '#,##0.00'
            ws.cell(row=row_num, column=col_idx + 2).value = netto_amount
            ws.cell(row=row_num, column=col_idx + 2).number_format = '#,##0.00'
            ws.cell(row=row_num, column=col_idx + 3).value = by_payment_method['gotowka']
            ws.cell(row=row_num, column=col_idx + 3).number_format = '#,##0.00'
            ws.cell(row=row_num, column=col_idx + 4).value = by_payment_method['karta']
            ws.cell(row=row_num, column=col_idx + 4).number_format = '#,##0.00'
            ws.cell(row=row_num, column=col_idx + 5).value = by_payment_method['blik']
            ws.cell(row=row_num, column=col_idx + 5).number_format = '#,##0.00'
            row_total += amount
            row_commission_total += commission_amount
            row_netto_total += netto_amount

            col_idx += 6

        total_col = col_idx
        ws.cell(row=row_num, column=total_col).value = row_total
        ws.cell(row=row_num, column=total_col).number_format = '#,##0.00'
        ws.cell(row=row_num, column=total_col + 1).value = carrier_cashless_total
        ws.cell(row=row_num, column=total_col + 1).number_format = '#,##0.00'

        prowizja_suma_col = total_col + 2
        dodatkowe_zyski_col = total_col + 3
        netto_suma_col = total_col + 4
        extra_ref = f"{get_column_letter(dodatkowe_zyski_col)}{row_num}"
        per_carrier_commission_refs = [
            f"{get_column_letter(4 + idx * 6 + 1)}{row_num}"
            for idx in range(len(carrier_order))
        ]
        sum_commission_expr = '+'.join(per_carrier_commission_refs) if per_carrier_commission_refs else '0'
        ws.cell(row=row_num, column=prowizja_suma_col).value = f"={sum_commission_expr}+IF({extra_ref}=\"\",0,{extra_ref})"
        ws.cell(row=row_num, column=prowizja_suma_col).number_format = '#,##0.00'
        ws.cell(row=row_num, column=dodatkowe_zyski_col).value = None
        ws.cell(row=row_num, column=dodatkowe_zyski_col).number_format = '#,##0.00'
        ws.cell(row=row_num, column=netto_suma_col).value = row_netto_total
        ws.cell(row=row_num, column=netto_suma_col).number_format = '#,##0.00'

        cost_entry = dict(costs_by_device.get(device_id, {}))
        it_card_value = 0.0
        for carrier_code, carrier_cashless in carrier_cashless_by_code.items():
            rate = IT_CARD_RATE_BY_CARRIER.get(carrier_code, DEFAULT_IT_CARD_RATE)
            it_card_value += carrier_cashless * rate
        cost_entry['it_card'] = it_card_value

        if automat_type == 'BB':
            bb_type_count += 1
            normalized_location = _normalize_text_for_match(lokalizacja)
            if normalized_location and normalized_location in bb_amort_sources['by_location']:
                cost_entry['amortyzacja'] = float(bb_amort_sources['by_location'][normalized_location] or 0.0)
                bb_matched_by_location += 1
            elif device_id in bb_amort_sources['by_nr_inw']:
                cost_entry['amortyzacja'] = float(bb_amort_sources['by_nr_inw'][device_id] or 0.0)
                bb_matched_by_nr += 1
            else:
                cost_entry['amortyzacja'] = 0.0
                bb_missing += 1

        costs_col = total_col + 5
        tvm_cost_sum = 0.0
        for key in TVM_COST_KEYS:
            val = float(cost_entry.get(key, 0.0) or 0.0)
            tvm_cost_sum += val
            ws.cell(row=row_num, column=costs_col).value = val
            ws.cell(row=row_num, column=costs_col).number_format = '#,##0.00'
            costs_col += 1

        ws.cell(row=row_num, column=costs_col).value = tvm_cost_sum
        ws.cell(row=row_num, column=costs_col).number_format = '#,##0.00'
        costs_col += 1

        other_costs_offset = len(OTHER_COST_KEYS)
        other_total = 0.0
        for offset, key in enumerate(OTHER_COST_KEYS):
            val = float(cost_entry.get(key, 0.0) or 0.0)
            other_total += val
            ws.cell(row=row_num, column=costs_col + offset).value = val
            ws.cell(row=row_num, column=costs_col + offset).number_format = '#,##0.00'

        all_costs_total = tvm_cost_sum + other_total

        ws.cell(row=row_num, column=costs_col + other_costs_offset).value = other_total
        ws.cell(row=row_num, column=costs_col + other_costs_offset).number_format = '#,##0.00'
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 1).value = all_costs_total
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 1).number_format = '#,##0.00'

        ws.cell(row=row_num, column=costs_col + other_costs_offset + 2).value = card_total
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 2).number_format = '#,##0.00'
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 3).value = blik_total
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 3).number_format = '#,##0.00'
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 4).value = cash_total
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 4).number_format = '#,##0.00'
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 5).value = card_total + blik_total
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 5).number_format = '#,##0.00'

        prowizja_ref = f"{get_column_letter(prowizja_suma_col)}{row_num}"
        suma_koszty_ref = f"{get_column_letter(costs_col + other_costs_offset + 1)}{row_num}"
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 6).value = f"={prowizja_ref}-{suma_koszty_ref}"
        ws.cell(row=row_num, column=costs_col + other_costs_offset + 6).number_format = '#,##0.00'
        ws.cell(row=row_num, column=uwagi_col_idx).value = None
        row_num += 1

    last_data_row = row_num - 1

    if last_data_row >= 2:
        summary_row = row_num
        ws.cell(row=summary_row, column=1).value = 'PODSUMOWANIE'
        ws.cell(row=summary_row, column=2).value = None
        ws.cell(row=summary_row, column=3).value = None
        ws.cell(row=summary_row, column=uwagi_col_idx).value = None

        for col_num in range(4, uwagi_col_idx):
            col_letter = get_column_letter(col_num)
            summary_cell = ws.cell(row=summary_row, column=col_num)
            summary_cell.value = f"=SUM({col_letter}2:{col_letter}{last_data_row})"
            summary_cell.number_format = '#,##0.00'

        for col_num in range(1, uwagi_col_idx + 1):
            cell = ws.cell(row=summary_row, column=col_num)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

        row_num += 1

    if bb_type_count:
        print(
            "✓ Amortyzacja BB: "
            f"typ BB={bb_type_count}, "
            f"po Nazwa={bb_matched_by_location}, "
            f"po Nr.inw={bb_matched_by_nr}, "
            f"braki={bb_missing}"
        )

    if last_data_row >= 2:
        result_col_letter = ws.cell(row=1, column=result_col_idx).column_letter
        result_range = f"{result_col_letter}2:{result_col_letter}{last_data_row}"
        ws.conditional_formatting.add(
            result_range,
            CellIsRule(
                operator='greaterThanOrEqual',
                formula=['0'],
                fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'),
            ),
        )
        ws.conditional_formatting.add(
            result_range,
            CellIsRule(
                operator='lessThan',
                formula=['0'],
                fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
            ),
        )
    
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
    print(f"✓ Liczba wierszy: {len(all_device_ids)}")


def _build_month_export_payload(conn, month_str, args, dictionary_comparison):
    """
    Buduje dane wejściowe do eksportu dla wskazanego miesiąca.
    """
    if args.obrot_source_mode == 'carrier-transactions':
        revenue_data = get_monthly_revenue_by_carrier(
            conn,
            month_str,
            carriers=args.obrot_carriers,
            table_name=args.obrot_transactions_table,
            amount_col=args.obrot_transactions_amount_col,
            date_col=args.obrot_transactions_date_col,
            device_col=args.obrot_transactions_device_col,
            payment_method_col=args.obrot_transactions_payment_method_col,
            use_device_range_filter=not args.obrot_no_device_range_filter,
        )
    else:
        revenue_data = get_monthly_revenue(
            conn,
            month_str,
            schema=args.schema,
            actioncodes=args.obrot_actioncodes,
            mctype=args.obrot_mctype,
            use_device_range_filter=not args.obrot_no_device_range_filter,
        )

    commission_rules, xlsx_location_by_device = load_commission_rules_and_locations_from_xlsx(
        month_str,
        commission_file=args.prowizja_file,
    )

    all_device_ids = sorted(set(
        [int(k) for k in dictionary_comparison.keys()] +
        [int(k) for k in revenue_data.keys()]
    ))

    location_by_device = {
        int(device_id): location
        for device_id, location in xlsx_location_by_device.items()
        if int(device_id) in all_device_ids
    }

    missing_location_ids = [device_id for device_id in all_device_ids if device_id not in location_by_device]
    if missing_location_ids:
        fallback_location_by_device = get_locations_by_carrier(
            conn,
            device_ids=missing_location_ids,
            carriers=args.obrot_carriers,
        )
        for device_id, location in fallback_location_by_device.items():
            if device_id not in location_by_device and location:
                location_by_device[device_id] = location

    automat_type_by_device = get_automat_type_by_device_from_av(
        conn,
        all_device_ids,
        schema_name='AV',
        table_name='dictionary',
        value_col='value',
        groupid_col='groupid',
    )

    return {
        'revenue_data': revenue_data,
        'commission_rules': commission_rules,
        'location_by_device': location_by_device,
        'automat_type_by_device': automat_type_by_device,
    }


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
        help='Źródło obrotu: per schemat przewoźnika z <carrier>.transactions lub klasycznie z <schema>.moneystats'
    )
    parser.add_argument(
        '--obrot-carriers',
        nargs='*',
        default=CARRIERS,
        help='Lista przewoźników/schematów do pobrania (np. ARP IC KD KML KW LKA PR SKM)'
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
        help='Kolumna kwoty w tabeli transakcji (domyślnie: fin_nalezn)'
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
        help='Kolumna numeru automatu (domyślnie: tvm_tvm_id)'
    )
    parser.add_argument(
        '--obrot-transactions-payment-method-col',
        type=str,
        default=TRANSACTIONS_SOURCE_CONFIG['payment_method_col'],
        help='Kolumna metody płatności (domyślnie: fin_spos_opl)'
    )
    parser.add_argument(
        '--obrot-print-sql',
        action='store_true',
        help='Wypisz zapytania SQL per przewoźnik/schemat dla wskazanego miesiąca'
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
        '--prowizja-file',
        type=str,
        default=str(DEFAULT_PROWIZJE_FILE),
        help='Ścieżka do pliku Excel z prowizjami i lokalizacjami (domyślnie: Prowizje_AB.xlsx)'
    )
    parser.add_argument(
        '--koszty-file',
        type=str,
        default=str(DEFAULT_COSTS_FILE),
        help='Ścieżka do pliku Excel z kosztami projektu'
    )
    parser.add_argument(
        '--najem-file',
        type=str,
        default=str(DEFAULT_RENT_FILE),
        help='Ścieżka do pliku Excel z czynszem i energią per automat'
    )
    parser.add_argument(
        '--najem-sheet',
        type=str,
        default=None,
        help='Nazwa arkusza z najmem (domyślnie: rok z --miesiac)'
    )
    parser.add_argument(
        '--amortyzacja-file',
        type=str,
        default=str(DEFAULT_AMORTYZACJA_FILE),
        help='Ścieżka do pliku z amortyzacją miesięczną automatów'
    )
    parser.add_argument(
        '--amortyzacja-sheet',
        type=str,
        default=DEFAULT_AMORTYZACJA_SHEET,
        help='Nazwa arkusza z amortyzacją BB (domyślnie: bb8)'
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

    normalized_carriers = []
    for carrier in args.obrot_carriers or []:
        code = _normalize_carrier_code(carrier)
        if not code:
            continue
        if code not in normalized_carriers:
            normalized_carriers.append(code)

    disallowed = sorted(code for code in normalized_carriers if code in DISALLOWED_CARRIERS)
    if disallowed:
        parser.error(
            f"Przewoźnik(i) niedozwoleni: {', '.join(disallowed)}. "
            f"Dozwolone wartości: {', '.join(CARRIERS)}"
        )
    args.obrot_carriers = normalized_carriers
    
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
            raise SystemExit(1)
    except psycopg2.Error as e:
        print(f"❌ Błąd połączenia z bazą: {e}")
        raise SystemExit(1)

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
    print(f"\n[1/6] Pobieranie słownika automatów ({source_schema}.dictionary)...")
    current_snapshot = get_dictionary_snapshot(conn, schema=source_schema)
    print(f"✓ Pobrano {len(current_snapshot)} automatów ze słownika")
    
    # Porównaj z poprzednim miesiącem
    print("\n[2/6] Porównanie z poprzednim miesiącem...")
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
            f"\n[3/6] Pobieranie obrotu za {month_str} "
            f"(per schemat przewoźnika: <carrier>.{args.obrot_transactions_table})..."
        )
        if args.obrot_no_device_range_filter:
            print("  Filtr zakresów urządzeń: WYŁĄCZONY")
        print(f"  Przewoźnicy: {', '.join(args.obrot_carriers)}")

        if args.obrot_print_sql:
            print("  SQL diagnostyczny per schemat:")
            for carrier in args.obrot_carriers:
                carrier_code = _normalize_carrier_code(carrier)
                carrier_table_name = _resolve_transactions_table_for_carrier(
                    carrier_code,
                    args.obrot_transactions_table,
                )
                carrier_device_col = _resolve_transactions_device_col_for_carrier(
                    carrier_code,
                    args.obrot_transactions_device_col,
                )
                carrier_amount_col = _resolve_transactions_amount_col_for_carrier(
                    carrier_code,
                    args.obrot_transactions_amount_col,
                )
                carrier_payment_method_col = _resolve_transactions_payment_method_col_for_carrier(
                    carrier_code,
                    args.obrot_transactions_payment_method_col,
                )
                sql_preview = build_transactions_debug_query(
                    schema_name=carrier_code,
                    month_str=month_str,
                    table_name=carrier_table_name or args.obrot_transactions_table,
                    amount_col=carrier_amount_col,
                    date_col=args.obrot_transactions_date_col,
                    device_col=_preview_column_name(carrier_device_col) or args.obrot_transactions_device_col,
                    payment_method_col=_preview_column_name(carrier_payment_method_col) or args.obrot_transactions_payment_method_col,
                )
                print(
                    f"\n--- {carrier} (schema: {carrier_code}, table: {carrier_table_name}) ---\n"
                    f"{sql_preview}\n"
                )

        revenue_data = get_monthly_revenue_by_carrier(
            conn,
            month_str,
            carriers=args.obrot_carriers,
            table_name=args.obrot_transactions_table,
            amount_col=args.obrot_transactions_amount_col,
            date_col=args.obrot_transactions_date_col,
            device_col=args.obrot_transactions_device_col,
            payment_method_col=args.obrot_transactions_payment_method_col,
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
        print(f"\n[3/6] Pobieranie obrotu za {month_str} ({source_schema}.moneystats)...")
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
    print(f"\n[4/6] Pobieranie prowizji i lokalizacji z pliku {args.prowizja_file}...")
    commission_rules, xlsx_location_by_device = load_commission_rules_and_locations_from_xlsx(
        month_str,
        commission_file=args.prowizja_file,
    )
    print(f"✓ Wczytano reguły prowizji: {len(commission_rules)} par (automat, przewoźnik)")

    commission_data = build_monthly_commission_data_from_rules(
        revenue_data,
        commission_rules,
        month_str,
    )
    total_commission = sum(v['prowizja_zl'] for v in commission_data.values())
    print(f"  Prowizja łączna (z pliku): {total_commission:,.2f} zł")

    if args.reference_provision_xls:
        print("  🔎 Porównanie danych prowizji vs plik referencyjny .xls...")
        reference_map = load_reference_provision_from_xls(args.reference_provision_xls)
        comparison = compare_commission_with_reference(commission_data, reference_map)
        if comparison is None:
            print("  ⚠ Pominięto porównanie (nie udało się odczytać pliku .xls)")
        else:
            print(f"  Brakujące w SQL: {len(comparison['missing_in_sql'])}")
            print(f"  Nadmiarowe w SQL: {len(comparison['extra_in_sql'])}")
            print(f"  Różnice kwot: {len(comparison['mismatches'])}")
    
    # === KROK 4: Lokalizacje automatów z pliku prowizji ===
    print(f"\n[5/6] Uzupełnianie lokalizacji automatów z pliku prowizji...")
    all_device_ids = sorted(set(
        [int(k) for k in dictionary_comparison.keys()] +
        [int(k) for k in revenue_data.keys()]
    ))
    location_by_device = {
        int(device_id): location
        for device_id, location in xlsx_location_by_device.items()
        if int(device_id) in all_device_ids
    }

    missing_location_ids = [device_id for device_id in all_device_ids if device_id not in location_by_device]
    fallback_location_by_device = {}
    if missing_location_ids:
        fallback_location_by_device = get_locations_by_carrier(
            conn,
            device_ids=missing_location_ids,
            carriers=args.obrot_carriers,
        )
        for device_id, location in fallback_location_by_device.items():
            if device_id not in location_by_device and location:
                location_by_device[device_id] = location

    print(
        f"✓ Uzupełniono lokalizacje dla {len(location_by_device)} / {len(all_device_ids)} automatów "
        f"(z pliku: {len(xlsx_location_by_device)}, fallback: {len(fallback_location_by_device)})"
    )

    automat_type_by_device = get_automat_type_by_device_from_av(
        conn,
        all_device_ids,
        schema_name='AV',
        table_name='dictionary',
        value_col='value',
        groupid_col='groupid',
    )
    print(f"✓ Uzupełniono typ automatu dla {len(automat_type_by_device)} / {len(all_device_ids)} automatów")

    # === KROK 5: Eksport do Excel ===
    print(f"\n[6/6] Eksport do Excel...")
    filename = build_output_filename(month_str, naming_mode=args.output_naming)
    export_to_excel_PL(
        dictionary_comparison,
        revenue_data,
        month_str,
        filename,
        commission_rules=commission_rules,
        carriers=args.obrot_carriers,
        location_by_device=location_by_device,
        automat_type_by_device=automat_type_by_device,
        costs_file=args.koszty_file,
        rent_file=args.najem_file,
        rent_sheet=args.najem_sheet,
        amortyzacja_file=args.amortyzacja_file,
        amortyzacja_sheet=args.amortyzacja_sheet,
    )

    report_year = int(month_str.split('-')[0])
    january_month_str = f"{report_year}-01"
    if january_month_str != month_str:
        print(f"\n[EXTRA] Generowanie osobnego skoroszytu dla stycznia ({january_month_str})...")
        january_payload = _build_month_export_payload(
            conn,
            january_month_str,
            args,
            dictionary_comparison,
        )
        january_filename = build_output_filename(january_month_str, naming_mode='default')
        export_to_excel_PL(
            dictionary_comparison,
            january_payload['revenue_data'],
            january_month_str,
            january_filename,
            commission_rules=january_payload['commission_rules'],
            carriers=args.obrot_carriers,
            location_by_device=january_payload['location_by_device'],
            automat_type_by_device=january_payload['automat_type_by_device'],
            costs_file=args.koszty_file,
            rent_file=args.najem_file,
            rent_sheet=args.najem_sheet,
            amortyzacja_file=args.amortyzacja_file,
            amortyzacja_sheet=args.amortyzacja_sheet,
        )
    else:
        print("\n[EXTRA] Pominieto dodatkowy skoroszyt styczniowy (raport bazowy jest juz za styczen).")

    conn.close()
    


if __name__ == "__main__":
    main()