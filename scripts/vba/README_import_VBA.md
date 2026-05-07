# Import VBA for TVM Search

Target workbook: `C:\Users\antek\Desktop\PL_TVM_2026-03.xlsm`

## 1. Import standard module
1. Open VBA editor (`Alt+F11`).
2. Right click project -> `Import File...`.
3. Import `modTvmWyszukiwarka.bas`.

## 2. Add worksheet event code
1. In VBA editor, find sheet `Wyszukiwarka` under `Microsoft Excel Objects`.
2. Open its code window.
3. Paste content from `Wyszukiwarka_Worksheet_Change.vba.txt`.

## 3. Save and trust macros
1. Save workbook (`Ctrl+S`).
2. Ensure macros are enabled when opening the workbook.

## 4. Usage
1. Fill `C8` (NR TVM), `F8/G8` (period 1), `F10/G10` (period 2).
2. Click button `Szukaj / Porownaj` or just edit fields (auto-run).
3. Table updates in `J6:N8` and chart title becomes `Dane automatu {NR_TVM}`.
