param(
    [string]$SourceWorkbookPath = 'C:\Users\antek\Desktop\PL_TVM_2026-03.xlsx',
    [string]$TargetWorkbookPath = 'C:\Users\antek\Desktop\PL_TVM_2026-03.xlsm'
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$vbaRoot = Join-Path $scriptRoot 'vba'
$standardModulePath = Join-Path $vbaRoot 'modTvmWyszukiwarka.bas'
$sheetEventPath = Join-Path $vbaRoot 'Wyszukiwarka_Worksheet_Change.vba.txt'

if (-not (Test-Path $SourceWorkbookPath) -and -not (Test-Path $TargetWorkbookPath)) {
    throw "Nie znaleziono pliku zrodlowego: $SourceWorkbookPath"
}
if (-not (Test-Path $standardModulePath)) {
    throw "Nie znaleziono modulu VBA: $standardModulePath"
}
if (-not (Test-Path $sheetEventPath)) {
    throw "Nie znaleziono kodu zdarzenia arkusza: $sheetEventPath"
}

$standardModuleCode = Get-Content -LiteralPath $standardModulePath -Raw -Encoding UTF8
$sheetEventCode = Get-Content -LiteralPath $sheetEventPath -Raw -Encoding UTF8
$standardModuleCode = $standardModuleCode -replace '^\s*Attribute VB_Name = ".*"\r?\n', ''

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $null

try {
    $openPath = if (Test-Path $TargetWorkbookPath) { $TargetWorkbookPath } else { $SourceWorkbookPath }
    $wb = $excel.Workbooks.Open($openPath)

    if ($wb.FileFormat -ne 52) {
        $wb.SaveAs($TargetWorkbookPath, 52)
    } elseif ($wb.FullName -ne $TargetWorkbookPath) {
        $wb.SaveAs($TargetWorkbookPath, 52)
    }

    $searchSheet = $null
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -eq 'Wyszukiwarka') {
            $searchSheet = $ws
            break
        }
    }

    if ($null -eq $searchSheet) {
        $searchSheet = $wb.Worksheets.Add()
        $searchSheet.Name = 'Wyszukiwarka'
    }

    $searchSheet.Range('B7').Value2 = 'WYBOR AUTOMATU'
    $searchSheet.Range('B8').Value2 = 'NR TVM:'
    $searchSheet.Range('E6').Value2 = 'WYBOR ROKU I MIESIACA'
    $searchSheet.Range('E8').Value2 = '1'
    $searchSheet.Range('E10').Value2 = '2'
    $searchSheet.Range('F7').Value2 = 'Rok'
    $searchSheet.Range('G7').Value2 = 'Miesiac'
    $searchSheet.Range('F9').Value2 = 'Rok'
    $searchSheet.Range('G9').Value2 = 'Miesiac'

    $searchSheet.Range('J6').Value2 = ''
    $searchSheet.Range('K6').Value2 = 'Sprzedaz brutto'
    $searchSheet.Range('L6').Value2 = 'Netto'
    $searchSheet.Range('M6').Value2 = 'Przychody'
    $searchSheet.Range('N6').Value2 = 'Koszty'
    $searchSheet.Range('J7').Value2 = 'Okres 1'
    $searchSheet.Range('J8').Value2 = 'Okres 2'

    try {
        $searchSheet.Range('K7:N8').NumberFormat = '#,##0.00'
    } catch {
        try {
            $searchSheet.Range('K7:N8').NumberFormatLocal = '# ##0,00'
        } catch {
            # Regional number formats vary between Excel installations.
        }
    }

    $months = @(
        'Styczen', 'Luty', 'Marzec', 'Kwiecien', 'Maj', 'Czerwiec',
        'Lipiec', 'Sierpien', 'Wrzesien', 'Pazdziernik', 'Listopad', 'Grudzien'
    )

    for ($i = 0; $i -lt $months.Count; $i++) {
        $searchSheet.Cells.Item($i + 1, 27).Value2 = $months[$i]
    }
    $searchSheet.Columns('AA').Hidden = $true

    foreach ($addr in @('C8')) {
        $cell = $searchSheet.Range($addr)
        try { $cell.Validation.Delete() } catch { }
        $cell.Validation.Add(1, 1, 1, '1000', '9999')
    }

    foreach ($addr in @('F8', 'F10')) {
        $cell = $searchSheet.Range($addr)
        try { $cell.Validation.Delete() } catch { }
        $cell.Validation.Add(1, 1, 1, '2000', '2100')
    }

    foreach ($addr in @('G8', 'G10')) {
        $cell = $searchSheet.Range($addr)
        try { $cell.Validation.Delete() } catch { }
        $cell.Validation.Add(3, 1, 1, '=Wyszukiwarka!$AA$1:$AA$12')
        $cell.Validation.InCellDropdown = $true
    }

    foreach ($shape in $searchSheet.Shapes) {
        if ($shape.Name -eq 'btnPorownajTVM') {
            $shape.Delete()
            break
        }
    }

    $button = $searchSheet.Shapes.AddShape(5, $searchSheet.Range('B12').Left, $searchSheet.Range('B12').Top, 150, 32)
    $button.Name = 'btnPorownajTVM'
    $button.TextFrame.Characters().Text = 'Szukaj / Porownaj'
    $button.OnAction = 'UruchomPorownanieTVM'

    $chartObj = $null
    foreach ($co in $searchSheet.ChartObjects()) {
        if ($co.Name -eq 'WykresDaneAutomatu') {
            $chartObj = $co
            break
        }
    }

    if ($null -eq $chartObj) {
        $chartObj = $searchSheet.ChartObjects().Add(
            $searchSheet.Range('J10').Left,
            $searchSheet.Range('J10').Top,
            620,
            320
        )
        $chartObj.Name = 'WykresDaneAutomatu'
    }

    $chart = $chartObj.Chart
    $chart.ChartType = 51
    $chart.SetSourceData($searchSheet.Range('J6:N8'), 1)
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = 'DANE Automatu'
    $chart.HasLegend = $true

    $searchSheet.Range('B7:N11').Font.Bold = $true
    $searchSheet.Columns('B:N').AutoFit() | Out-Null

    $vbProject = $null
    try {
        $vbProject = $wb.VBProject
    } catch {
        $vbProject = $null
    }

    if ($null -eq $vbProject) {
        throw (
            "Brak dostepu do VBProject (Trust Center blokuje dostep programowy). " +
            "UI zostalo przygotowane, ale kod VBA trzeba zaimportowac recznie z: " +
            "$standardModulePath oraz $sheetEventPath"
        )
    }

    $existingModule = $null
    foreach ($component in $vbProject.VBComponents) {
        if ($component.Name -eq 'modTvmWyszukiwarka') {
            $existingModule = $component
            break
        }
    }

    if ($null -ne $existingModule) {
        $vbProject.VBComponents.Remove($existingModule)
    }

    $stdModule = $vbProject.VBComponents.Add(1)
    $stdModule.Name = 'modTvmWyszukiwarka'
    $stdModule.CodeModule.AddFromString($standardModuleCode)

    $sheetComponent = $vbProject.VBComponents.Item($searchSheet.CodeName)
    $lineCount = $sheetComponent.CodeModule.CountOfLines
    if ($lineCount -gt 0) {
        $sheetComponent.CodeModule.DeleteLines(1, $lineCount)
    }
    $sheetComponent.CodeModule.AddFromString($sheetEventCode)

    $wb.Save()
    Write-Output "OK: Makro wdrozone. Zapisano: $($wb.FullName)"
} finally {
    if ($wb) {
        $wb.Close($true)
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
