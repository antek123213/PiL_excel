Attribute VB_Name = "modTvmWyszukiwarka"
Option Explicit

Private Const SEARCH_SHEET_NAME As String = "Wyszukiwarka"
Private Const CHART_OBJECT_NAME As String = "WykresDaneAutomatu"

Public Sub UruchomPorownanieTVM(Optional ByVal CalledFromEvent As Boolean = False)
    On Error GoTo ErrHandler

    Dim wsSearch As Worksheet
    Set wsSearch = ThisWorkbook.Worksheets(SEARCH_SHEET_NAME)

    Dim tvmId As Long
    tvmId = ParseLong(wsSearch.Range("C8").Value)
    If tvmId <= 0 Then
        RaiseUserError "Pole C8 (NR TVM) musi byc poprawna liczba."
    End If

    Dim year1 As Long, year2 As Long
    year1 = ParseLong(wsSearch.Range("F8").Value)
    year2 = ParseLong(wsSearch.Range("F10").Value)

    ValidateYear year1, "F8"
    ValidateYear year2, "F10"

    Dim period1 As String, period2 As String
    period1 = BuildPeriodCode(year1, CStr(wsSearch.Range("G8").Value), "G8")
    period2 = BuildPeriodCode(year2, CStr(wsSearch.Range("G10").Value), "G10")

    Dim formulas1 As Variant
    Dim formulas2 As Variant
    formulas1 = LoadPeriodFormulas(tvmId, period1)
    formulas2 = LoadPeriodFormulas(tvmId, period2)

    PopulateComparisonTable wsSearch, period1, period2, formulas1, formulas2
    UpsertChart wsSearch, tvmId

    If Not CalledFromEvent Then
        MsgBox "Zaktualizowano dane dla automatu " & CStr(tvmId) & ".", vbInformation
    End If

    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Porownanie TVM"
End Sub

Private Function BuildPeriodCode(ByVal yearValue As Long, ByVal monthName As String, ByVal cellAddress As String) As String
    Dim monthNo As String
    monthNo = MonthNameToNumber(monthName)
    If monthNo = "" Then RaiseUserError "Nieprawidlowy miesiac w komorce " & cellAddress & "."

    BuildPeriodCode = CStr(yearValue) & "-" & monthNo
End Function

Private Function LoadPeriodFormulas(ByVal tvmId As Long, ByVal periodCode As String) As Variant
    Dim wsPeriod As Worksheet
    Set wsPeriod = GetPeriodWorksheet(periodCode)
    If wsPeriod Is Nothing Then
        RaiseUserError "Brak arkusza: P&L " & periodCode
    End If

    Dim rowNo As Long
    rowNo = FindTvmRow(wsPeriod, tvmId)
    If rowNo = 0 Then
        RaiseUserError "Nie znaleziono TVM " & CStr(tvmId) & " w arkuszu P&L " & periodCode & "."
    End If

    LoadPeriodFormulas = Array( _
        BuildFormulaForHeader(wsPeriod, rowNo, "Brutto Suma", "sprzedaz brutto"), _
        BuildFormulaForHeader(wsPeriod, rowNo, "Netto Suma", "netto"), _
        BuildRevenueFormula(wsPeriod, rowNo), _
        BuildFormulaForHeader(wsPeriod, rowNo, "Koszty", "koszty") _
    )
End Function

Private Function BuildRevenueFormula(ByVal wsPeriod As Worksheet, ByVal rowNo As Long) As String
    Dim przychodyCol As Long
    przychodyCol = FindHeaderColumn(wsPeriod, "Przychody")
    If przychodyCol > 0 Then
        BuildRevenueFormula = FormulaForCell(wsPeriod, rowNo, przychodyCol)
        Exit Function
    End If

    Dim parts As Collection
    Set parts = New Collection

    AddFormulaPart parts, wsPeriod, rowNo, "Prowizja Suma"
    AddFormulaPart parts, wsPeriod, rowNo, "Interchange"
    AddFormulaPart parts, wsPeriod, rowNo, "Dodatkowe zyski"

    If parts.Count = 0 Then
        RaiseUserError "Brak kolumny Przychody oraz skladnikow przychodow w arkuszu " & wsPeriod.Name & "."
    End If

    Dim formulaText As String
    Dim i As Long
    For i = 1 To parts.Count
        If formulaText <> "" Then formulaText = formulaText & "+"
        formulaText = formulaText & CStr(parts.Item(i))
    Next i

    BuildRevenueFormula = "=" & formulaText
End Function

Private Sub AddFormulaPart(ByVal parts As Collection, ByVal wsPeriod As Worksheet, ByVal rowNo As Long, ByVal headerText As String)
    Dim colNo As Long
    colNo = FindHeaderColumn(wsPeriod, headerText)
    If colNo > 0 Then
        parts.Add "'" & EscapeSheetName(wsPeriod.Name) & "'!" & wsPeriod.Cells(rowNo, colNo).Address(False, False)
    End If
End Sub

Private Function BuildFormulaForHeader(ByVal wsPeriod As Worksheet, ByVal rowNo As Long, ByVal headerText As String, ByVal labelForError As String) As String
    Dim colNo As Long
    colNo = FindHeaderColumn(wsPeriod, headerText)
    If colNo = 0 Then
        RaiseUserError "Brak kolumny " & labelForError & " (" & headerText & ") w arkuszu " & wsPeriod.Name & "."
    End If

    BuildFormulaForHeader = FormulaForCell(wsPeriod, rowNo, colNo)
End Function

Private Function FormulaForCell(ByVal wsPeriod As Worksheet, ByVal rowNo As Long, ByVal colNo As Long) As String
    FormulaForCell = "='" & EscapeSheetName(wsPeriod.Name) & "'!" & wsPeriod.Cells(rowNo, colNo).Address(False, False)
End Function

Private Function EscapeSheetName(ByVal sheetName As String) As String
    EscapeSheetName = Replace(sheetName, "'", "''")
End Function

Private Sub PopulateComparisonTable( _
    ByVal wsSearch As Worksheet, _
    ByVal period1 As String, _
    ByVal period2 As String, _
    ByVal formulas1 As Variant, _
    ByVal formulas2 As Variant _
)
    wsSearch.Range("J6:N8").ClearContents

    wsSearch.Range("J6").Value = ""
    wsSearch.Range("K6").Value = "Sprzedaz brutto"
    wsSearch.Range("L6").Value = "Netto"
    wsSearch.Range("M6").Value = "Przychody"
    wsSearch.Range("N6").Value = "Koszty"

    wsSearch.Range("J7").Value = period1
    wsSearch.Range("J8").Value = period2

    wsSearch.Range("K7").Formula = formulas1(0)
    wsSearch.Range("L7").Formula = formulas1(1)
    wsSearch.Range("M7").Formula = formulas1(2)
    wsSearch.Range("N7").Formula = formulas1(3)

    wsSearch.Range("K8").Formula = formulas2(0)
    wsSearch.Range("L8").Formula = formulas2(1)
    wsSearch.Range("M8").Formula = formulas2(2)
    wsSearch.Range("N8").Formula = formulas2(3)

    wsSearch.Range("K7:N8").NumberFormat = "#,##0.00"
    wsSearch.Columns("J:N").AutoFit
End Sub

Private Sub UpsertChart(ByVal wsSearch As Worksheet, ByVal tvmId As Long)
    Dim chartObj As ChartObject

    On Error Resume Next
    Set chartObj = wsSearch.ChartObjects(CHART_OBJECT_NAME)
    On Error GoTo 0

    If chartObj Is Nothing Then
        Set chartObj = wsSearch.ChartObjects.Add( _
            Left:=wsSearch.Range("J10").Left, _
            Top:=wsSearch.Range("J10").Top, _
            Width:=620, _
            Height:=320 _
        )
        chartObj.Name = CHART_OBJECT_NAME
    End If

    With chartObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=wsSearch.Range("J6:N8"), PlotBy:=xlRows
        .HasTitle = True
        .ChartTitle.Text = "DANE Automatu " & CStr(tvmId)
        .HasLegend = True
    End With
End Sub

Private Function GetPeriodWorksheet(ByVal periodCode As String) As Worksheet
    On Error Resume Next
    Set GetPeriodWorksheet = ThisWorkbook.Worksheets("P&L " & periodCode)
    On Error GoTo 0
End Function

Private Function FindTvmRow(ByVal ws As Worksheet, ByVal tvmId As Long) As Long
    Dim tvmCol As Long
    tvmCol = FindHeaderColumn(ws, "Nr TVM")
    If tvmCol = 0 Then tvmCol = FindHeaderColumn(ws, "Nr aut.")
    If tvmCol = 0 Then tvmCol = 2

    Dim found As Range
    Set found = ws.Columns(tvmCol).Find(What:=tvmId, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
    If found Is Nothing Then
        Set found = ws.Columns(tvmCol).Find(What:=CStr(tvmId), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
    End If

    If Not found Is Nothing Then
        FindTvmRow = found.Row
    Else
        FindTvmRow = 0
    End If
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim target As String
    target = NormalizeText(headerText)

    Dim rowNo As Long
    Dim lastCol As Long
    Dim colNo As Long
    Dim valueText As String

    For rowNo = 1 To 5
        lastCol = ws.Cells(rowNo, ws.Columns.Count).End(xlToLeft).Column
        For colNo = 1 To lastCol
            valueText = NormalizeText(CStr(ws.Cells(rowNo, colNo).Value))
            If valueText = target Then
                FindHeaderColumn = colNo
                Exit Function
            End If
        Next colNo
    Next rowNo

    FindHeaderColumn = 0
End Function

Private Function ParseLong(ByVal value As Variant) As Long
    If IsNumeric(value) Then
        ParseLong = CLng(value)
    Else
        ParseLong = 0
    End If
End Function

Private Sub ValidateYear(ByVal yearValue As Long, ByVal cellAddress As String)
    If yearValue < 2000 Or yearValue > 2100 Then
        RaiseUserError "Nieprawidlowy rok w komorce " & cellAddress & "."
    End If
End Sub

Private Function MonthNameToNumber(ByVal monthName As String) As String
    Select Case NormalizeText(monthName)
        Case "STYCZEN": MonthNameToNumber = "01"
        Case "LUTY": MonthNameToNumber = "02"
        Case "MARZEC": MonthNameToNumber = "03"
        Case "KWIECIEN": MonthNameToNumber = "04"
        Case "MAJ": MonthNameToNumber = "05"
        Case "CZERWIEC": MonthNameToNumber = "06"
        Case "LIPIEC": MonthNameToNumber = "07"
        Case "SIERPIEN": MonthNameToNumber = "08"
        Case "WRZESIEN": MonthNameToNumber = "09"
        Case "PAZDZIERNIK": MonthNameToNumber = "10"
        Case "LISTOPAD": MonthNameToNumber = "11"
        Case "GRUDZIEN": MonthNameToNumber = "12"
        Case Else: MonthNameToNumber = ""
    End Select
End Function

Private Function NormalizeText(ByVal value As String) As String
    Dim s As String
    s = UCase$(Trim$(value))

    s = Replace(s, ChrW(260), "A")
    s = Replace(s, ChrW(262), "C")
    s = Replace(s, ChrW(280), "E")
    s = Replace(s, ChrW(321), "L")
    s = Replace(s, ChrW(323), "N")
    s = Replace(s, ChrW(211), "O")
    s = Replace(s, ChrW(346), "S")
    s = Replace(s, ChrW(377), "Z")
    s = Replace(s, ChrW(379), "Z")

    s = Replace(s, ChrW(261), "A")
    s = Replace(s, ChrW(263), "C")
    s = Replace(s, ChrW(281), "E")
    s = Replace(s, ChrW(322), "L")
    s = Replace(s, ChrW(324), "N")
    s = Replace(s, ChrW(243), "O")
    s = Replace(s, ChrW(347), "S")
    s = Replace(s, ChrW(378), "Z")
    s = Replace(s, ChrW(380), "Z")

    NormalizeText = s
End Function

Private Sub RaiseUserError(ByVal messageText As String)
    Err.Raise vbObjectError + 513, "UruchomPorownanieTVM", messageText
End Sub
