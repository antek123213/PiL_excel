## Plan: Miesięczne zestawienie TVM wg schematu 1-8

Plan jest oparty bezpośrednio na Twojej „TABELI GŁÓWNEJ" (kolumny 1-8) i zakłada: raport za pełny poprzedni miesiąc, uruchamianie automatyczne + ręczne, koszty tymczasowo z Excel/CSV i docelowo z DB. Kolumny `2-6` są danymi miesięcznymi (zmienne m/m), a `1-INFO` jest pobierane co miesiąc w trybie kontroli zmian względem poprzedniego miesiąca.

**Steps**
1. Faza A: Zamrożenie struktury raportu wg Twojego szkicu.
2. Ustalić finalny układ tabeli wyjściowej z dokładnym mapowaniem kolumn 1-8: `1-INFO`, `2-OBRÓT`, `3-PROWIZJA`, `4-Rodzaj tranzakcji`, `5-KOSZTY`, `6-PO SUMIE`, `7-PODSUMOWANIE ROCZNE`, `8-UWAGI` (*blokuje dalsze kroki*).
3. Wprowadzić podział logiczny zgodny z Twoją decyzją: `1-INFO` pobierane raz na miesiąc z porównaniem do poprzedniego miesiąca i aktualizacją tylko przy zmianie, natomiast `2-6` zawsze liczone/zasilane dla konkretnego miesiąca raportowego (*depends on 2*).Aby pobierac dane z bazy danych w DB należy wejść do katalogu `AV` i tabeli `dictionary` i tam są 4 kolumny. 1 to jest id automatu jako klucz głowny tabeli i masz description, czyli opis automatu, który jest unikalny. Tylko gdy raz zrobisz raport dla danego automatu, to wtedy możesz porównywać z poprzednim miesiącem. W przypadku gdy raportujesz po raz pierwszy, to wtedy nie masz poprzedniego miesiąca i wtedy po prostu pobierasz dane i zapisujesz jako pierwszy snapshot. Przy kolejnych raportach porównujesz z tym snapshotem i aktualizujesz tylko przy zmianie. Dla kolumn `2-6` zawsze pobierasz dane dla konkretnego miesiąca raportowego, niezależnie od tego czy jest to pierwszy raport czy kolejny.
4. Faza B: Warstwa danych źródłowych.
5. Zdefiniować źródła i klucze łączenia po `nr_automatu`, `rok`, `miesiąc`, przewoźniku: `1-INFO` jako miesięczny snapshot z kontrolą zmian vs poprzedni miesiąc, `2-OBRÓT , 3-Prowizja , 4-Rodzaj tranzakcji` jako regularne miesięczne pobranie z bazy, `5-KOSZTY` wg źródeł finansowych i kosztowych (*depends on 2*).
6. Dla kosztów zewnętrznych (wg szkicu) przygotować osobny strumień danych i mapowanie do kolumny `5-KOSZTY` (*parallel with 5*).
7. Dodać regułę okresu raportowego: domyślnie pełny poprzedni miesiąc + tryb ręczny `YYYY-MM` (*parallel with 5,6*).
8. Faza C: Kalkulacje wg kolumn 3-6.
9. Zaimplementować kalkulacje prowizji `3-PROWIZJA` per przewoźnik i operator płatności (ELAVON/interchange, zgodnie ze szkicem) (*depends on 5*).
10. Zaimplementować klasyfikację `4-RODZAJ TRANZAKCJI` na podstawie danych z DB i reguł biznesowych z Twojej notatki (*depends on 5*).
11. Zaimplementować agregację kosztów `5-KOSZTY` z rozbiciem na pozycje ze szkicu (czynsz, prąd, ELAVON, poczta, amortyzacja, serwis, transmisja i pozostałe pozycje po potwierdzeniu), wszystkie koszty bedą podawane recznie z excela (*depends on 6*).
12. Zbudować `6-PO SUMIE` jako wynik netto po wszystkich składowych z kolumn 2-5 + tagi segmentów (np. OM/Nord/marketing po potwierdzeniu) (*depends on 9,10,11*).
13. Faza D: Widoki raportowe 7 i 8.
14. Dodać `7-PODSUMOWANIE ROCZNE`: suma roczna obrotu, prowizji, kosztów i wyniku, z możliwością filtrowania po roku i automacie (*depends on 12*).
15. Dodać `8-UWAGI` jako pole manualne/narracyjne o najniższym priorytecie (nie blokuje wyliczeń finansowych) (*parallel with 14*).
16. Pobierac automaty numery od '1101' do '1299'
16. Faza E: Eksport, harmonogram, dystrybucja.
17. Przygotować eksport do Excela w układzie zgodnym z tabelą główną (*depends on 14,15*).
18. Dodać uruchamianie automatyczne (Windows Task Scheduler, 1. dzień miesiąca, 01:00) oraz ręczne parametrem miesiąca (*depends on 17*).
19. Dodać powiadomienie e-mail/Teams z KPI: obrót, prowizja, koszty, wynik netto i status wykonania (*depends on 17*).
20. Faza F: Migracja kosztów do DB.
21. Utrzymać adapter źródeł kosztów (Excel/CSV -> DB) tak, by przełączenie odbyło się konfiguracją bez zmiany logiki kolumn 3-6 (*depends on 11*).
22. Wykonać porównanie 2-3 miesięcy (wynik z Excel vs wynik z DB) i po zgodności przełączyć domyślne źródło na DB (*depends on 21*).

**Relevant files**
- `c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\P&L\export_automaty.py` - główna implementacja pobierania, kalkulacji i eksportu.
- `c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\P&L\P&L.txt` - reguły biznesowe pozycji P&L do mapowania kolumn 3-6.
- `c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\P&L\kas.txt` - podstawa dla danych czynszowych i kosztów zewnętrznych.
- `c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\P&L\config.(ini|yaml|json)` - konfiguracja źródeł danych, harmonogramu i powiadomień.
- `c:\Users\antek\OneDrive - University of Gdansk (for Students)\Dokumenty\Praca\P&L\scheduler\task_scheduler.md` - instrukcja uruchamiania cyklicznego.

**Verification**
1. Walidacja struktury: raport zawiera wszystkie sekcje 1-8 w kolejności ze szkicu.
2. Walidacja logiki: `1-INFO` jest co miesiąc porównywane z poprzednim snapshotem i aktualizowane tylko przy różnicach, natomiast `2-6` zmieniają się m/m zgodnie z danymi dla wskazanego miesiąca raportowego.
3. Walidacja finansowa: porównanie 1-2 miesięcy z ręcznym plikiem kontrolnym.
4. Walidacja automatyzacji: wykonanie z harmonogramu + wykonanie ręczne dla `YYYY-MM` daje identyczny wynik.
5. Walidacja dystrybucji: poprawny plik Excel/CSV i wysłana notyfikacja e-mail/Teams.

**Decisions**
- Raport bazuje na „TABELI GŁÓWNEJ" 1-8 z Twojej notatki.
- `1-INFO` jest pobierane raz na miesiąc i porównywane z poprzednim miesiącem; aktualizacja tylko przy zmianie.
- `2-6` są miesięczne i zmienne, liczone każdorazowo dla wybranego miesiąca (`2-OBRÓT` z bazy danych co miesiąc).
- Okres domyślny: pełny poprzedni miesiąc.
- Start: koszty z Excel/CSV; cel: koszty z DB.
- Dystrybucja: Excel + notyfikacja e-mail/Teams.
- In scope: automatyzacja miesięcznego raportu finansowego TVM.
- Out of scope: dashboard BI i analityka predykcyjna.

**Further Considerations**
1. Do potwierdzenia: dokładne definicje kolumn `4-PODATEK` i `6-PO SUMIE` (w szkicu część skrótów jest niejednoznaczna).
2. Do potwierdzenia: pełna lista pozycji w `5-KOSZTY` (część pozycji w notatce jest skrótowa/nieczytelna).
3. Rekomendacja: zamrozić słownik nazw kolumn i pozycji kosztowych przed implementacją, żeby uniknąć zmian formatu po starcie automatyzacji.
