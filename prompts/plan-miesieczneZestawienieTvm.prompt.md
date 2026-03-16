## Plan: Miesięczne zestawienie TVM wg schematu 1-8

Plan jest oparty bezpośrednio na Twojej „TABELI GŁÓWNEJ" i zakłada: raport za pełny poprzedni miesiąc, uruchamianie automatyczne + ręczne, koszty tymczasowo z Excel/CSV i docelowo z DB. W aktualnym układzie raportu pozostają sekcje `1-INFO`, `2-OBRÓT`, `3-PROWIZJA`, `5-KOSZTY`, `7-PODSUMOWANIE ROCZNE`, `8-UWAGI` (kolumny 4 i 6 usunięte). Kolumny miesięczne to `2,3,5`, a `1-INFO` jest pobierane co miesiąc w trybie kontroli zmian względem poprzedniego miesiąca.

**Steps**
1. Faza A: Zamrożenie struktury raportu wg Twojego szkicu.
2. Ustalić finalny układ tabeli wyjściowej z dokładnym mapowaniem sekcji: `1-INFO`, `2-OBRÓT`, `3-PROWIZJA`, `5-KOSZTY`, `7-PODSUMOWANIE ROCZNE`, `8-UWAGI` (*blokuje dalsze kroki*).
3. Wprowadzić podział logiczny zgodny z Twoją decyzją: `1-INFO` pobierane raz na miesiąc z porównaniem do poprzedniego miesiąca i aktualizacją tylko przy zmianie, natomiast `2,3,5` zawsze liczone/zasilane dla konkretnego miesiąca raportowego (*depends on 2*).Aby pobierac dane z bazy danych w DB należy wejść do katalogu `AV` i tabeli `dictionary` i tam są 4 kolumny. 1 to jest id automatu jako klucz głowny tabeli i masz description, czyli opis automatu, który jest unikalny. Tylko gdy raz zrobisz raport dla danego automatu, to wtedy możesz porównywać z poprzednim miesiącem. W przypadku gdy raportujesz po raz pierwszy, to wtedy nie masz poprzedniego miesiąca i wtedy po prostu pobierasz dane i zapisujesz jako pierwszy snapshot. Przy kolejnych raportach porównujesz z tym snapshotem i aktualizujesz tylko przy zmianie. Dla kolumn `2,3,5` zawsze pobierasz dane dla konkretnego miesiąca raportowego, niezależnie od tego czy jest to pierwszy raport czy kolejny.
4. Faza B: Warstwa danych źródłowych.
5. Zdefiniować źródła i klucze łączenia po `nr_automatu`, `rok`, `miesiąc`, przewoźniku: `1-INFO` jako miesięczny snapshot z kontrolą zmian vs poprzedni miesiąc, `2-OBRÓT , 3-Prowizja` jako regularne miesięczne pobranie z bazy, `5-KOSZTY` wg źródeł finansowych i kosztowych (*depends on 2*).
6. Dla kosztów zewnętrznych (wg szkicu) przygotować osobny strumień danych i mapowanie do kolumny `5-KOSZTY` (*parallel with 5*).
7. Dodać regułę okresu raportowego: domyślnie pełny poprzedni miesiąc + tryb ręczny `YYYY-MM` (*parallel with 5,6*).
8. Faza C: Kalkulacje wg kolumn 3 i 5.
9. Zaimplementować kalkulacje prowizji `3-PROWIZJA` per przewoźnik i operator płatności (ELAVON/interchange, zgodnie ze szkicem) (*depends on 5*).
11. Zaimplementować agregację kosztów `5-KOSZTY` z rozbiciem na pozycje ze szkicu (czynsz, prąd, ELAVON, poczta, amortyzacja, serwis, transmisja i pozostałe pozycje po potwierdzeniu), wszystkie koszty bedą podawane recznie z excela (*depends on 6*).
12. Wynik netto liczyć pomocniczo poza układem kolumn głównych (kolumna 6 usunięta) i wykorzystać go tylko do KPI/podsumowań technicznych (*depends on 9,11*).
13. Faza D: Widoki raportowe 7 i 8.
14. Dodać `7-PODSUMOWANIE ROCZNE`: suma roczna obrotu, prowizji i kosztów (oraz ewentualnego wyniku pomocniczego), z możliwością filtrowania po roku i automacie (*depends on 12*).
15. Dodać `8-UWAGI` jako pole manualne/narracyjne o najniższym priorytecie (nie blokuje wyliczeń finansowych) (*parallel with 14*).
16. Pobierac automaty z zakresów: '1101-1141' oraz '1201-1299'.
17. Faza E: Eksport, harmonogram, dystrybucja.
18. Przygotować eksport do Excela w układzie zgodnym z tabelą główną (*depends on 14,15*).
19. Dodać uruchamianie automatyczne (Windows Task Scheduler, 1. dzień miesiąca, 01:00) oraz ręczne parametrem miesiąca (*depends on 18*).
20. Dodać powiadomienie e-mail/Teams z KPI: obrót, prowizja, koszty, wynik netto i status wykonania (*depends on 18*).
21. Faza F: Migracja kosztów do DB.
22. Utrzymać adapter źródeł kosztów (Excel/CSV -> DB) tak, by przełączenie odbyło się konfiguracją bez zmiany logiki kolumn 3 i 5 (*depends on 11*).
23. Wykonać porównanie 2-3 miesięcy (wynik z Excel vs wynik z DB) i po zgodności przełączyć domyślne źródło na DB (*depends on 22*).

**Relevant files**
- `c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\P&L\export_automaty.py` - główna implementacja pobierania, kalkulacji i eksportu.
- `c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\P&L\P&L.txt` - reguły biznesowe pozycji P&L do mapowania kolumn 3-6.
- `c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\P&L\kas.txt` - podstawa dla danych czynszowych i kosztów zewnętrznych.
- `c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\P&L\config.(ini|yaml|json)` - konfiguracja źródeł danych, harmonogramu i powiadomień.
- `c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\P&L\scheduler\task_scheduler.md` - instrukcja uruchamiania cyklicznego.

**Verification**
1. Walidacja struktury: raport zawiera sekcje `1,2,3,5,7,8` (kolumny 4 i 6 nie występują).
2. Walidacja logiki: `1-INFO` jest co miesiąc porównywane z poprzednim snapshotem i aktualizowane tylko przy różnicach, natomiast `2,3,5` zmieniają się m/m zgodnie z danymi dla wskazanego miesiąca raportowego.
3. Walidacja finansowa: porównanie 1-2 miesięcy z ręcznym plikiem kontrolnym.
4. Walidacja automatyzacji: wykonanie z harmonogramu + wykonanie ręczne dla `YYYY-MM` daje identyczny wynik.
5. Walidacja dystrybucji: poprawny plik Excel/CSV i wysłana notyfikacja e-mail/Teams.

**Decisions**
- Raport bazuje na „TABELI GŁÓWNEJ" po uproszczeniu do sekcji 1,2,3,5,7,8 (kolumny 4 i 6 usunięte).
- `1-INFO` jest pobierane raz na miesiąc i porównywane z poprzednim miesiącem; aktualizacja tylko przy zmianie.
- `2,3,5` są miesięczne i zmienne, liczone każdorazowo dla wybranego miesiąca (`2-OBRÓT` z bazy danych co miesiąc).
- Zakres automatów: 1101-1141 oraz 1201-1299.
- Okres domyślny: pełny poprzedni miesiąc.
- Start: koszty z Excel/CSV; cel: koszty z DB.
- Dystrybucja: Excel + notyfikacja e-mail/Teams.
- In scope: automatyzacja miesięcznego raportu finansowego TVM.
- Out of scope: dashboard BI i analityka predykcyjna.

**Further Considerations**
1. Do potwierdzenia: czy wynik netto ma występować wyłącznie jako KPI/podsumowanie techniczne (bez osobnej kolumny w tabeli).
2. Do potwierdzenia: pełna lista pozycji w `5-KOSZTY` (część pozycji w notatce jest skrótowa/nieczytelna).
3. Rekomendacja: zamrozić słownik nazw kolumn i pozycji kosztowych przed implementacją, żeby uniknąć zmian formatu po starcie automatyzacji.
