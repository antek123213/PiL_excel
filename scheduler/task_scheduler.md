# Automatyczne uruchamianie raportu TVM P&L – Windows Task Scheduler

## Cel

Raport TVM P&L generowany jest automatycznie 1. dnia każdego miesiąca o godzinie 01:00,
obejmując pełny poprzedni miesiąc. Możliwe jest też ręczne uruchomienie dla dowolnego
okresu YYYY-MM.

---

## 1. Automatyczne uruchamianie (Windows Task Scheduler)

### 1.1 Tworzenie zadania przez GUI (zalecane)

1. Otwórz **Harmonogram zadań** (`taskschd.msc`).
2. Kliknij **Utwórz zadanie…** (nie Utwórz zadanie podstawowe).
3. Zakładka **Ogólne**:
   - Nazwa: `TVM_PL_Miesieczny`
   - Opis: `Miesięczne zestawienie TVM P&L wg schematu 1-8`
   - Uruchom jako: konto z dostępem do bazy danych i katalogu projektu
   - Zaznacz: **Uruchom niezależnie od tego, czy użytkownik jest zalogowany**
   - Zaznacz: **Uruchom z najwyższymi uprawnieniami**
4. Zakładka **Wyzwalacze** → Nowy wyzwalacz:
   - Rozpocznij: **Miesięcznie**
   - Dzień: **1**
   - Godz.: **01:00:00**
   - Zaznacz: **Włączone**
5. Zakładka **Akcje** → Nowa akcja:
   - Akcja: **Uruchom program**
   - Program: `C:\Python311\python.exe`
   - Argumenty: `C:\Praca\PiL_excel\export_automaty.py --config C:\Praca\PiL_excel\config.local.yaml`
   - Rozpocznij w: `C:\Praca\PiL_excel`
6. Zakładka **Warunki**:
   - Odznacz "Uruchamiaj zadanie tylko wtedy, gdy komputer jest zasilany z sieci" (opcjonalne).
7. Zakładka **Ustawienia**:
   - Zaznacz: **Uruchom zadanie tak szybko jak to możliwe po pominięciu uruchomienia zaplanowanego**
8. Kliknij **OK** i podaj hasło konta.

---

### 1.2 Tworzenie zadania przez CLI (schtasks)

Uruchom w wierszu polecenia z uprawnieniami administratora:

```cmd
schtasks /create ^
  /tn "TVM_PL_Miesieczny" ^
  /tr "C:\Python311\python.exe C:\Praca\PiL_excel\export_automaty.py --config C:\Praca\PiL_excel\config.local.yaml" ^
  /sc MONTHLY /d 1 /st 01:00 ^
  /ru SYSTEM ^
  /rl HIGHEST ^
  /f
```

> **Uwaga:** Zamień ścieżki na rzeczywiste ścieżki instalacji Pythona i projektu.

---

### 1.3 Weryfikacja zadania

```cmd
schtasks /query /tn "TVM_PL_Miesieczny" /fo LIST /v
```

Ręczne uruchomienie testowe:

```cmd
schtasks /run /tn "TVM_PL_Miesieczny"
```

---

## 2. Ręczne uruchamianie

### 2.1 Poprzedni miesiąc (domyślny)

```cmd
cd C:\Praca\PiL_excel
python export_automaty.py
```

### 2.2 Wybrany miesiąc (YYYY-MM)

```cmd
cd C:\Praca\PiL_excel
python export_automaty.py --month 2026-02
```

### 2.3 Z alternatywną konfiguracją

```cmd
python export_automaty.py --month 2026-02 --config config.local.yaml
```

---

## 3. Środowisko i zależności

### 3.1 Instalacja zależności Python

```cmd
pip install -r requirements.txt
```

### 3.2 Zmienne środowiskowe (zamiast haseł w config)

| Zmienna          | Opis                                  |
|------------------|---------------------------------------|
| `DB_PASSWORD`    | Hasło do bazy danych                  |
| `EMAIL_PASSWORD` | Hasło do konta e-mail (SMTP)          |

Ustaw je systemowo lub w pliku `.env` (używaj python-dotenv).

---

## 4. Weryfikacja poprawności wykonania

Po każdym uruchomieniu (automatycznym lub ręcznym) sprawdź:

1. **Plik Excel** – `output\TVM_PL_YYYY_MM.xlsx` – powinien istnieć i zawierać dane.
2. **Plik CSV** – `output\TVM_PL_YYYY_MM_audit.csv` – plik audytowy.
3. **Logi** – standardowe wyjście procesu (przechwytywane przez Task Scheduler do Event Log).
4. **Powiadomienie** – e-mail / Teams z KPI.
5. **Snapshot INFO** – `snapshots\info_snapshot_YYYY_MM.json` – plik z danymi 1-INFO.

---

## 5. Obsługa błędów

| Scenariusz                    | Zachowanie                                              |
|-------------------------------|---------------------------------------------------------|
| Brak połączenia z DB          | Błąd w logu, status e-mail/Teams: `BŁĄD: ...`          |
| Brak pliku kosztów (CSV/xlsx) | Ostrzeżenie w logu, koszty = 0, uwaga w kolumnie 8     |
| Brak danych obrotu dla automatu | Ostrzeżenie, wiersz z zerowym obrotem + uwaga w kol. 8 |
| Błąd wysyłania e-mail         | Błąd w logu, raport Excel/CSV wygenerowany              |

---

## 6. Walidacja zgodności wyników (ręczna vs automatyczna)

Aby potwierdzić identyczność wyników:

```cmd
python export_automaty.py --month 2026-01
python export_automaty.py --month 2026-01
```

Oba uruchomienia dla tego samego miesiąca powinny generować identyczny plik Excel/CSV
(deterministyczne obliczenia, ta sama baza danych).
