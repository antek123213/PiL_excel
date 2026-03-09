import psycopg2
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime

# Konfiguracja połączenia z bazą PostgreSQL
DB_CONFIG = {
    'host': '192.168.101.20',
    'port': 5432,
    'database': 'monitor',
    'user': 'sprzedaz',
    'password': 'tVregNm5',
    'schema': 'POZNAŃ'        # zmień schemat jeśli trzeba
}

def get_automaty_data():
    try:
        conn = psycopg2.connect(
            host=DB_CONFIG['host'],
            port=DB_CONFIG['port'],
            database=DB_CONFIG['database'],
            user=DB_CONFIG['user'],
            password=DB_CONFIG['password']
        )
        cursor = conn.cursor()
        # Zapytanie SQL - pobiera automaty z numerami 1101-1299 ze WSZYSTKICH schematów (przewoźników)
        # Łączy dane z tabeli transactions ze wszystkich dostępnych przewoźników
        query = """
        SELECT DISTINCT
            tvm_tvm_id AS "ID Automatu",
            tvm_automatnum AS "Numer Automatu",
            CASE WHEN position(':' in tvm_description) > 0
                THEN substring(tvm_description from 1 for (position(':' in tvm_description)-1))
                ELSE tvm_description 
            END AS "Lokalizacja"
        FROM (
            SELECT tvm_tvm_id, tvm_automatnum, tvm_description FROM "POZNAŃ".transactions
            UNION ALL
            SELECT tvm_tvm_id, tvm_automatnum, tvm_description FROM "KRAKÓW".transactions
            UNION ALL
            SELECT tvm_tvm_id, tvm_automatnum, tvm_description FROM "GDAŃSK".transactions
            UNION ALL
            SELECT tvm_tvm_id, tvm_automatnum, tvm_description FROM "WROCŁAW".transactions
            UNION ALL
            SELECT tvm_tvm_id, tvm_automatnum, tvm_description FROM "SZCZECIN".transactions
            UNION ALL
            SELECT tvm_tvm_id, tvm_automatnum, tvm_description FROM "IC".transactions
            UNION ALL
            SELECT tvm_tvm_id, tvm_automatnum, tvm_description FROM "PR".transactions
        ) AS all_automaty
        WHERE tvm_automatnum >= 1101 
          AND tvm_automatnum <= 1299
          AND tvm_tvm_id IS NOT NULL
        ORDER BY tvm_automatnum
        """
        
        cursor.execute(query)
        columns = [desc[0] for desc in cursor.description]
        data = cursor.fetchall()
        
        cursor.close()
        conn.close()
        
        return columns, data
        
    except psycopg2.Error as e:
        print(f"Błąd połączenia z bazą: {e}")
        return None, None

def export_to_excel(columns, data, filename='automaty.xlsx'):
    """Eksportuje dane do pliku Excel"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Automaty"
    
    # Nagłówki
    for col_num, col_title in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = col_title
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    # Dane
    for row_num, row_data in enumerate(data, 2):
        for col_num, value in enumerate(row_data, 1):
            ws.cell(row=row_num, column=col_num).value = value
    
    # Dostosowanie szerokości kolumn
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Zapis pliku
    wb.save(filename)
    print(f"✓ Dane eksportowane do: {filename}")
    print(f"✓ Liczba wierszy: {len(data)}")

def main():
    print("Pobieranie danych o automatach...")
    columns, data = get_automaty_data()
    
    if data is None:
        print("Nie udało się pobrać danych")
        return
    
    if not data:
        print("Brak danych w bazie")
        return
    
    print(f"Znaleziono {len(data)} automatów")
    
    # Ścieżka do pliku Excel
    excel_file = r'c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\automaty.xlsx'
    
    export_to_excel(columns, data, excel_file)

if __name__ == "__main__":
    main()
