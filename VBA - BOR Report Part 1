Sub PART_ONE()
'
' PART_ONE Macro
'
' Keyboard Shortcut: Ctrl+Shift+O
'
    Application.ScreenUpdating = False
    
    'PREP THE RAW DATA
    Dim Sheet1Range As Range
    Dim Sheet1RangeCount  As Long
    On Error Resume Next
    Range("A2").Select

    Set Sheet1Range = Range(Selection, Selection.End(xlDown))
    If Sheet1Range Is Nothing Then Resume Next
    Sheet1RangeCount = Sheet1Range.Count
    
    'Replace Empty RFOs with Description Updates
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""",RC[3],RC[-1])"
    Range("P2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 16), Cells(Sheet1RangeCount + 1, 16))
    Range(Cells(2, 16), Cells(Sheet1RangeCount + 1, 16)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("O2").Select
    ActiveSheet.Paste
    Range("P2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("O2").Select
    
    'Replace empty Impact couts with 1
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""",1,RC[-1])"
    Range("K2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 11), Cells(Sheet1RangeCount + 1, 11))
    Range(Cells(2, 11), Cells(Sheet1RangeCount + 1, 11)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Impact Count"
    '____________________________________________________________________________________________
    'Remove cancelled tickets
    Selection.AutoFilter
    ActiveSheet.Range("A1").AutoFilter Field:=20, Criteria1:=Array( _
        "Closed", "Resolved", "WIP"), Operator:=xlFilterValues
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add().Name = "Temp"
    Sheets("Temp").Select
    ActiveSheet.Paste
    Sheets(2).Select
    Application.CutCopyMode = False
    Range("A1").Select
    Selection.AutoFilter
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Sheets("Temp").Select
    Selection.Copy
    Sheets(2).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Temp").Select
    'Application.DisplayAlerts = False
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    'Application.DisplayAlerts = True
    Sheets(2).Select
    Range("A1").Select
    
    '_______________________________________________________________________________________
    'Rearrange the Columns
    
    Columns("T:T").Select
    Selection.Cut
    Columns("S:S").Select
    Selection.Insert Shift:=xlToRight
    Application.CutCopyMode = False
    Columns("H:H").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("I:I").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("K:K").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Columns("I:I").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("R:R").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Columns("P:P").Select
    Selection.Cut
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight
    Columns("S:S").Select
    Selection.Cut
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight
    Columns("O:O").Select
    Selection.Cut
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight
    Columns("N:N").Select
    Selection.ClearContents
    Range("O1").Select
    Columns("R:R").Select
    Selection.Cut
    Columns("P:P").Select
    Selection.Insert Shift:=xlToRight
    Columns("R:R").Select
    Selection.Cut
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight
    Columns("S:S").Select
    Selection.Cut
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Ticket ID"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Issue"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "NE ID"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Region"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Name of Affected Node"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "ID of Affected Node"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "NO of NE's Affected"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Outage Date and Time"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Fault Recovery Date and Time"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Duration"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Severity"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Root Cause"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Affected Service"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Status (Open/Closed)"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Sub Network Category"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "RFO"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Power Vendor"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Instability Occurrence"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .VerticalAlignment = xlTop
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    '_________________________________________________________________________________
    'Replace Column N with Business Status and replace WIP with Open, Resolved with Closed
    Range("N2").Select
    Sheets(2).Name = "Sheet1"
    Columns("T:T").Select
    Selection.Cut
    Columns("N:N").Select
    ActiveSheet.Paste
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Status (Open/Closed)"
    Columns("N:N").Select
    Selection.Replace What:="WIP", Replacement:="Open", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Resolved", Replacement:="Closed", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("T:T").Select
    Selection.Delete Shift:=xlLeft
    Range("A1").Select
    
    'Sort Headers to get opened tickets at the top
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range(Cells(2, 14), Cells(Sheet1RangeCount, 14)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range(Cells(2, 8), Cells(Sheet1RangeCount, 8)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range(Cells(1, 1), Cells(Sheet1RangeCount, 20))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    '_____________________________________________________________________________________
    'Get new column S to help get longest running open ticket
    Range("S2").Select
    Columns("S:S").Select
    Selection.Copy
    Columns("U:U").Select
    ActiveSheet.Paste
    
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "=TODAY()-RC[-11]"
    Range("S2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 19), Cells(Sheet1RangeCount, 19))
    Range(Cells(2, 19), Cells(Sheet1RangeCount, 19)).Select
    Selection.NumberFormat = "@"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("S2").Select
    
    '______________________________________________________________________________________
    'First get the lenght of the column to know where to stop when using autofill
    'Range("A" & Rows.Count).End(xlUp).Offset(0, 15).Select
    'Enter the formular
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-1]-RC[-2])"
    'Apply autofill
    Selection.AutoFill Destination:=Range(Cells(2, 10), Cells(Sheet1RangeCount, 10))
    'Change format to hh:mm, Copy and paste values
    Range(Cells(2, 10), Cells(Sheet1RangeCount, 10)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "hh:mm"
    Range("A1").Select
    
    '________________________________________________________________________________
    'Removing Tickets from previous days
    
    Dim Title As String
    
    Dim newYear, newMonth, newDay As String
    newYear = Year(Date)
    newMonth = Month(Date)
    newDay = Day(Date)
    
    If Len(newMonth) < 2 Then
    newMonth = "0" & newMonth
    End If
    If newDay >= 2 Then
    newDay = newDay - 1
    End If
    If Len(newDay) < 2 Then
    newDay = "0" & newDay
    End If
    
    Title = "INC-" & newYear & newMonth & newDay
    
    'ActiveSheet.Range("A1").AutoFilter Field:=14, Criteria1:="=*Open*" _
        , Operator:=xlAnd
    'Debug.Print Title
    'Range("A2").Select
    'Range(Cells(2, 1), Cells(2, 21)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.Copy
    'Sheets("Temp").Select
    'ActiveSheet.Paste
    'Range("A1").Select
    'Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    'Sheets("Sheet1").Select
    'Application.CutCopyMode = False
    'Application.DisplayAlerts = False
    'Selection.EntireRow.Delete
    'Application.DisplayAlerts = True
    'ActiveSheet.Range("A1").AutoFilter Field:=14
    'Range("A1").Select
    'ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:= _
    '    Title & "*", Operator:=xlAnd
    'Range("A2").Select
    'Range(Cells(2, 1), Cells(2, 21)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.Copy
    'Sheets("Temp").Select
    'ActiveSheet.Paste
    'Sheets("Sheet1").Select
    'Application.CutCopyMode = False
    'Application.DisplayAlerts = False
    'Selection.EntireRow.Delete
    'Application.DisplayAlerts = True
    'ActiveSheet.Range("A1").AutoFilter Field:=1
    'Range("A2").Select
    'Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    'Selection.Delete Shift:=xlToLeft
    'Range("A2").Select
    'Sheets("Temp").Select
    'Range("A1").Select
    'Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    'Selection.Cut
    'Sheets("Sheet1").Select
    'ActiveSheet.Paste
    'Range("A1").Select
    'Selection.AutoFilter
    'Sheets("Temp").Select
    'Range("A1").Select
    'Sheets("Sheet1").Select
    
    'Add new sheet and populate it
    Range("A1").Select
    Range(Cells(1, 1), Cells(1, 20)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Dim iMonth, iYear, iDay As Long
    Dim MyMonth, MyDate As String

    iMonth = Month(Date)
    iYear = Year(Date)
    iDay = Day(Date) - 1
    MyMonth = MonthName(iMonth, False)
    MyDate = "Business Outage Report for " & iDay & " " & MyMonth & ", " & iYear
    
    Application.DisplayAlerts = False
    Set thisWb = ActiveWorkbook
    Workbooks.Add
    Range("A1").Select
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Filename:=thisWb.Path & "\" & MyDate & ".xlsx"
    Range("A1").Select
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    'ActiveWorkbook.Close savechanges:=False
    
    'CONTINUE WITH PART ONE
    
    '1. CREATE THE SHEETS

    Sheets.Add().Name = "PENDING"
    Sheets.Add().Name = "Total"
    Sheets.Add().Name = "Enterprise"
    Sheets.Add().Name = "Telecom"
    Sheets.Add().Name = "RAN"
    Sheets.Add().Name = "TX"
    Sheets.Add().Name = "CORE Network"
    Sheets.Add().Name = "MPLS-IN-VAS"
    Sheets.Add().Name = "SUMMARY"
    
    
    '2. DRAW THE SUMMARY TABLES
    Sheets("SUMMARY").Select
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "MPLS/IN/VAS"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Count"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Outage minutes"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Sites down before and Fixed on this date"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Sites which were down before and still down"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "Sites went down and recovered"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "Sites went down and slipped to next day"
    Range("B1:C1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
    End With
    Range("D1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
    End With
    Range("B1:D5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Selection.Copy
    Range("B11").Select
    ActiveSheet.Paste
    Range("B21").Select
    ActiveSheet.Paste
    Range("B31").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B11").Select
    ActiveCell.FormulaR1C1 = "Core Network"
    Range("B21").Select
    ActiveCell.FormulaR1C1 = "RAN Network"
    Range("B31").Select
    ActiveCell.FormulaR1C1 = "TX Network"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "MPLS/IN/VAS"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "COUNT/DURATION"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "REF TT"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "Total Tickets"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "Open Tickets"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "Oldest Open Ticket from When"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "Closed Tickets"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "Average MTTR for Closed"
    Range("F7").Select
    ActiveCell.FormulaR1C1 = "Fastest Closed Ticket Time"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Longest Closed Ticket Time"
    Range("F1:I1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
    End With
    Range("F1:I8").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Selection.Copy
    Range("F11").Select
    ActiveSheet.Paste
    Range("F21").Select
    ActiveSheet.Paste
    Range("F31").Select
    ActiveSheet.Paste
    Range("F41").Select
    ActiveSheet.Paste
    Range("F51").Select
    ActiveSheet.Paste
    Range("F61").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("F11").Select
    ActiveCell.FormulaR1C1 = "Core Network"
    Range("F21").Select
    ActiveCell.FormulaR1C1 = "RAN Network"
    Range("F31").Select
    ActiveCell.FormulaR1C1 = "TX Network"
    Range("F41").Select
    ActiveCell.FormulaR1C1 = "Telecome"
    Range("F41").Select
    ActiveCell.FormulaR1C1 = "Telecom"
    Range("F51").Select
    ActiveCell.FormulaR1C1 = "Enterprise"
    Range("F61").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "TELECOM"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "ENTERPRISE"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "MPLS-IN-VAS"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "CORE"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "TRANSMISSION"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "RAN"
    Range("K6").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    
    Range("L1:N1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
    End With
    
    Range("K1:N6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Range("K8").Select
    ActiveCell.FormulaR1C1 = "Total Tickets"
    Range("L9").Select
    ActiveCell.FormulaR1C1 = "P1"
    Range("M9").Select
    ActiveCell.FormulaR1C1 = "P2"
    Range("N9").Select
    ActiveCell.FormulaR1C1 = "P3"
    Range("O9").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("K10").Select
    ActiveCell.FormulaR1C1 = "MPLS-IN-VAS"
    Range("K11").Select
    ActiveCell.FormulaR1C1 = "CORE"
    Range("K12").Select
    ActiveCell.FormulaR1C1 = "TRANSMISSION"
    Range("K13").Select
    ActiveCell.FormulaR1C1 = "RAN"
    Range("K14").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    
    Range("K9:O14").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("L9:O9").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
    End With
    
    Range("K9:O14").Select
    Selection.Copy
    Range("K17").Select
    ActiveSheet.Paste
    Range("K25").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("K8").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
    End With
    
    Range("K16").Select
    Selection.FormulaR1C1 = "Closed TTs"
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("K24").Select
    Selection.FormulaR1C1 = "Open Tickets"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
    End With
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    'FILL N/A CELLS
    'MPLS
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H6").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    'Core
    Range("H12").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H13").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H15").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H16").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    'RAN
    Range("H22").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H23").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H25").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H26").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    'TX
    Range("H32").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H33").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H36").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    'Telecom
    Range("H42").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H43").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H45").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H46").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    'Enterprise
    Range("H52").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H53").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H55").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H56").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    'Total
    Range("H62").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H63").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H65").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H66").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("A1").Select
    
    
    '____________________________________________________________________________________________
    
    '3. POPULATE THE FIRST 4 SHEETS AND PENDING SHEET
    Dim TestVar As String
    Dim TestRange As Range
    
    '3a FILTER FOR DATA
    Sheets("Sheet1").Select
    Range("A1").Select
    Range("A1").AutoFilter Field:=2
    Range("A1").AutoFilter Field:=2, Criteria1:="=*DATA:*" _
        , Operator:=xlAnd
    Range("A1").Select
    
    Set TestRange = Range(Selection, Selection.End(xlDown))
    'Set TestRange = TestRange(1)
    TestVar = Cells(TestRange(1).End(xlDown).Row, TestRange.Column)
    
    If TestVar = "" Then
    Range("A1").AutoFilter Field:=2
    Range("A1:R1").Select
    Selection.Copy
    Sheets("MPLS-IN-VAS").Select
    Range("A1").Select
    ActiveSheet.Paste
    Else:
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets("MPLS-IN-VAS").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    End If
    
    
    '3b FILTER FOR CORE
    Sheets("Sheet1").Select
    Range("A1").AutoFilter Field:=2
    Range("A1").AutoFilter Field:=2, Criteria1:="=*CORE:*" _
        , Operator:=xlAnd
    Range("A1").Select
    
    Set TestRange = Range(Selection, Selection.End(xlDown))
    TestVar = Cells(TestRange(1).End(xlDown).Row, TestRange.Column).Value
    
    If TestVar = "" Then
    Range("A1").AutoFilter Field:=2
    Range("A1:R1").Select
    Selection.Copy
    Sheets("CORE Network").Select
    Range("A1").Select
    ActiveSheet.Paste
    Else:
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets("CORE Network").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    End If
    
   '3c FILTER FOR TX
    Sheets("Sheet1").Select
    Range("A1").AutoFilter Field:=2
    Range("A1").AutoFilter Field:=2, Criteria1:="=*TX:*" _
        , Operator:=xlAnd
    Range("A1").Select
    
    Set TestRange = Range(Selection, Selection.End(xlDown))
    TestVar = Cells(TestRange(1).End(xlDown).Row, TestRange.Column).Value
    
    If TestVar = "" Then
    Range("A1").AutoFilter Field:=2
    Range("A1:R1").Select
    Selection.Copy
    Sheets("TX").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    Else:
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets("TX").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    End If
    
    
    '3d FILTER FOR RAN
    Sheets("Sheet1").Select
    Range("A1").AutoFilter Field:=2
    Range("A1").AutoFilter Field:=2, Criteria1:="=*RAN:*" _
        , Operator:=xlAnd
    Range("A1").Select
    
    Set TestRange = Range(Selection, Selection.End(xlDown))
    TestVar = Cells(TestRange(1).End(xlDown).Row, TestRange.Column).Value
    
    If TestVar = "" Then
    Range("A1").AutoFilter Field:=2
    Range("A1:R1").Select
    Selection.Copy
    Sheets("RAN").Select
    Range("A1").Select
    ActiveSheet.Paste
    Else:
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets("RAN").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    End If
    
    '________________________________________________________________________________________
    
    '4a. PREP THE RAN TABLE TO ADD IMPACTED SITES
    Sheets("RAN").Select
    Dim xRg As Range
    Dim I, xNum, xLastRow, xFstRow, xCol, xCount, ColSeclet, ColumnMerge  As Long
    On Error Resume Next
    Range("G2").Select

    Set xRg = Range(Selection, Selection.End(xlDown))
    If xRg Is Nothing Then Resume Next
    xLastRow = xRg(1).End(xlDown).Row
    xFstRow = xRg.Row
    xCol = xRg.Column
    xCount = xRg.Count
    Set xRg = xRg(1)
    For I = xLastRow To xFstRow Step -1
        xNum = Cells(I, xCol)
        ' -1 has been added to the value assigned to xNum to make sure the resize includes the main row
        If IsNumeric(xNum) And xNum > 1 Then
            Rows(I + 1).Resize(xNum - 1).Insert
            xCount = xCount + xNum
            
            ' Enter code block to merge added rows
            ColSeclet = I + Cells(I, xCol) - 1
            ColumnMerge = 1
            While (ColumnMerge <= 4)
            Range(Cells(I, ColumnMerge), Cells(ColSeclet, ColumnMerge)).Select
            Selection.Merge
            ColumnMerge = ColumnMerge + 1
            Wend
    
            ColumnMerge = 7
            While (ColumnMerge <= 19)
            Range(Cells(I, ColumnMerge), Cells(ColSeclet, ColumnMerge)).Select
            Selection.Merge
            ColumnMerge = ColumnMerge + 1
            Wend
            ' End of Merging code block
            
        End If
    Next
    
        
    '4b. PREP THE TX TABLE TO ADD IMPACTED SITES
    Sheets("TX").Select
    On Error Resume Next
    Range("G2").Select
    
    Set xRg = Range(Selection, Selection.End(xlDown))
    If xRg Is Nothing Then Resume Next
    xLastRow = xRg(1).End(xlDown).Row
    xFstRow = xRg.Row
    xCol = xRg.Column
    xCount = xRg.Count
    Set xRg = xRg(1)
    For I = xLastRow To xFstRow Step -1
        xNum = Cells(I, xCol)
       ' -1 has been added to the value assigned to xNum to make sure the resize includes the main row
        If IsNumeric(xNum) And xNum > 1 Then
            Rows(I + 1).Resize(xNum - 1).Insert
            xCount = xCount + xNum
    
            ' Enter code block to merge added rows
            ColSeclet = I + Cells(I, xCol) - 1
            ColumnMerge = 1
            While (ColumnMerge <= 4)
            Range(Cells(I, ColumnMerge), Cells(ColSeclet, ColumnMerge)).Select
            Selection.Merge
            ColumnMerge = ColumnMerge + 1
            Wend
    
            ColumnMerge = 7
            While (ColumnMerge <= 19)
            Range(Cells(I, ColumnMerge), Cells(ColSeclet, ColumnMerge)).Select
            Selection.Merge
            ColumnMerge = ColumnMerge + 1
            Wend
            ' End of Merging code block
    
        End If
    Next

    
    
    Sheets("Sheet1").Select
    Selection.AutoFilter
    Range("A1").Select
    Sheets("RAN").Select
    
    
    
    Application.ScreenUpdating = True

'
End Sub
