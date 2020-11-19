Sub PART_TWO()
'
' PART_TWO Macro
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    Application.ScreenUpdating = False
    
    '4c. Get the site IDs of the sites pasted
    Dim ShtRANRng As Range
    Dim ShtRANRngCount  As Long
    Sheets("RAN").Select
    Range("E2").Select

    Set ShtRANRng = Range(Selection, Selection.End(xlDown))
    If ShtRANRng Is Nothing Then Resume Next
    ShtRANRngCount = ShtRANRng.Count

    
    Range("F2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'[VDF BOR GENERATOR - 3.2.xlsm]RAN_Impact_Lookup'!C1:C2,2,FALSE)"
    Range("F2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 6), Cells(ShtRANRngCount + 1, 6))
    Range(Cells(2, 6), Cells(ShtRANRngCount + 1, 6)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A1").Select
    
    
    '4d. Get the site IDs of the sites pasted in TX
    Sheets("TX").Select
    Range("E2").Select

    Set ShtRANRng = Range(Selection, Selection.End(xlDown))
    If ShtRANRng Is Nothing Then Resume Next
    ShtRANRngCount = ShtRANRng.Count
    
    Range("F2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'[VDF BOR GENERATOR - 3.2.xlsm]RAN_Impact_Lookup'!C1:C2,2,FALSE)"
    Range("F2").Select
    Selection.AutoFill Destination:=Range(Cells(2, 6), Cells(ShtRANRngCount + 1, 6))
    Range(Cells(2, 6), Cells(ShtRANRngCount + 1, 6)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A1").Select
    
    
    '5. POPULATE TELECOM SHEET
    
    '5a. Copy from MPLS Sheet
    Sheets("MPLS-IN-VAS").Select
    Range("A2").Select
    If ActiveCell.Value = "" Then
    Range("A1").Select
    GoTo CoreSheet
    Else
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("Telecom").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    Sheets("MPLS-IN-VAS").Select
    Range("A1").Select
    Application.CutCopyMode = False
    End If
    
CoreSheet:
    Sheets("CORE Network").Select
    Range("A2").Select
    If ActiveCell.Value = "" Then
    Range("A1").Select
    GoTo TXSheet
    Else
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("Telecom").Select
    Range("A1").Select
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    ActiveSheet.Paste
    Range("A1").Select
    Sheets("CORE Network").Select
    Range("A1").Select
    Application.CutCopyMode = False
    End If
    
TXSheet:
    Sheets("TX").Select
    Range("A2").Select
    If ActiveCell.Value = "" Then
    Range("A1").Select
    GoTo RANSheet
    Else
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("Telecom").Select
    Range("A1").Select
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    ActiveSheet.Paste
    Range("A1").Select
    Sheets("TX").Select
    Range("A1").Select
    Application.CutCopyMode = False
    End If
    
RANSheet:
    Sheets("RAN").Select
    Range("A2").Select
    If ActiveCell.Value = "" Then
    Range("A1").Select
    GoTo StartEnterprice
    Else
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("Telecom").Select
    Range("A1").Select
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    ActiveSheet.Paste
    Range("A1").Select
    Sheets("RAN").Select
    Range("A1").Select
    Application.CutCopyMode = False
    End If
    
    'Start Populating Enterprise Sheet
StartEnterprice:
    Sheets("Telecom").Select
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("Enterprise").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    Sheets("Total").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    Sheets("Telecom").Select
    Range("A1").Select
    
    
    'Remove CORPORATE issues from Telecom Sheet
    Dim a, b As Integer

    Sheets("Telecom").Select
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    ActiveCell.Offset(0, 14).Range("A1").Select
    a = ActiveCell.Row
    While a > 1
    If ActiveCell.Value = "CORPORATE" Or ActiveCell.Value = "Sub Network Category" Then
    Rows(a).Select
    Selection.Delete Shift:=xlUp
    Cells(a, 15).Select
    ActiveCell.Offset(-1, 0).Range("A1").Select
    a = ActiveCell.Row
    Else
    ActiveCell.Offset(-1, 0).Range("A1").Select
    a = ActiveCell.Row
    End If
    Wend
    
    'Enterprise Sheet
    'Remove TELECOM from Enterprise Sheet
    Sheets("Enterprise").Select
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    ActiveCell.Offset(0, 14).Range("A1").Select
    a = ActiveCell.Row
    While a > 1
    If ActiveCell.Value = "TELECOM" Or ActiveCell.Value = "Sub Network Category" Then
    Rows(a).Select
    Selection.Delete Shift:=xlUp
    Cells(a, 15).Select
    ActiveCell.Offset(-1, 0).Range("A1").Select
    a = ActiveCell.Row
    Else
    ActiveCell.Offset(-1, 0).Range("A1").Select
    a = ActiveCell.Row
    End If
    Wend

    
    'Total Sheet
    'Remove TELECOM and CORPORATE from Total Sheet
    Sheets("Total").Select
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    ActiveCell.Offset(0, 14).Range("A1").Select
    a = ActiveCell.Row
    While a > 1
    If ActiveCell.Value = "Sub Network Category" Then
    Rows(a).Select
    Selection.Delete Shift:=xlUp
    Cells(a, 15).Select
    ActiveCell.Offset(-1, 0).Range("A1").Select
    a = ActiveCell.Row
    Else
    ActiveCell.Offset(-1, 0).Range("A1").Select
    a = ActiveCell.Row
    End If
    Wend
    
    
    '______________________________________________________________________________________
    '7. POPULATE SUMMARY TABLES
    '7a. MPLS/IN/VAS
    Sheets("SUMMARY").Select
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('MPLS-IN-VAS'!C[-6])-1"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('MPLS-IN-VAS'!C[7], ""Open"")"
    
    'To get the longest running openned ticket, filter for openned tickets and return the largest
    If Range("G3").Value = 0 Then
    GoTo Label1
    End If
    Sheets("MPLS-IN-VAS").Select
    Range("A1").AutoFilter Field:=14, Criteria1:="=*Open*", Operator:=xlAnd
    Range("H" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("G4").Select
    ActiveSheet.Paste
    Sheets("MPLS-IN-VAS").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("H4").Select
    ActiveSheet.Paste
    Sheets("MPLS-IN-VAS").Select
    Range("A" & Rows.Count).End(xlUp).Offset(0, 19).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("I4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("MPLS-IN-VAS").Select
    Selection.AutoFilter
    Sheets("SUMMARY").Select
    
Label1:
    Sheets("SUMMARY").Select
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('MPLS-IN-VAS'!C[7], ""Closed"")"
    
    'If total tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G2").Value = 0 Then
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H7").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I7").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages1
    End If
    
    'If closed tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G5").Value = 0 Then
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H7").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I7").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages1
    End If
    
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(AVERAGE('MPLS-IN-VAS'!C[3]))=0,MINUTE(AVERAGE('MPLS-IN-VAS'!C[3]))&""mins"",IF(HOUR(AVERAGE('MPLS-IN-VAS'!C[3]))>1,HOUR(AVERAGE('MPLS-IN-VAS'!C[3]))&""hrs, ""&MINUTE(AVERAGE('MPLS-IN-VAS'!C[3]))&""mins"",HOUR(AVERAGE('MPLS-IN-VAS'!C[3]))&""hr, ""&MINUTE(AVERAGE('MPLS-IN-VAS'!C[3]))&""mins""))"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SMALL('MPLS-IN-VAS'!C[3],1))=0,MINUTE(SMALL('MPLS-IN-VAS'!C[3],1))&""mins"",IF(HOUR(SMALL('MPLS-IN-VAS'!C[3],1))>1,HOUR(SMALL('MPLS-IN-VAS'!C[3],1))&""hrs, ""&MINUTE(SMALL('MPLS-IN-VAS'!C[3],1))&""mins"",HOUR(SMALL('MPLS-IN-VAS'!C[3],1))&""hr, ""&MINUTE(SMALL('MPLS-IN-VAS'!C[3],1))&""mins""))"
    Range("G8").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(LARGE('MPLS-IN-VAS'!C[3],1))=0,MINUTE(LARGE('MPLS-IN-VAS'!C[3],1))&""mins"",IF(HOUR(LARGE('MPLS-IN-VAS'!C[3],1))>1,HOUR(LARGE('MPLS-IN-VAS'!C[3],1))&""hrs, ""&MINUTE(LARGE('MPLS-IN-VAS'!C[3],1))&""mins"",HOUR(LARGE('MPLS-IN-VAS'!C[3],1))&""hr, ""&MINUTE(LARGE('MPLS-IN-VAS'!C[3],1))&""mins""))"
    Range("H7").Select
    ActiveCell.FormulaR1C1 = "=INDEX('MPLS-IN-VAS'!C[-7], MATCH(SMALL('MPLS-IN-VAS'!C[2],1),'MPLS-IN-VAS'!C[2],0))"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "=INDEX('MPLS-IN-VAS'!C[-7], MATCH(LARGE('MPLS-IN-VAS'!C[2],1), 'MPLS-IN-VAS'!C[2],0))"
    Range("I7").Select
    ActiveCell.FormulaR1C1 = "=INDEX('MPLS-IN-VAS'!C[7], MATCH(SMALL('MPLS-IN-VAS'!C[1],1), 'MPLS-IN-VAS'!C[1],0))"
    Range("I8").Select
    ActiveCell.FormulaR1C1 = "=INDEX('MPLS-IN-VAS'!C[7], MATCH(LARGE('MPLS-IN-VAS'!C[1],1), 'MPLS-IN-VAS'!C[1],0))"
    
SkipAverages1:
    '7b. Core Network
    Range("G12").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('Core Network'!C[-6])-1"
    Range("G13").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('Core Network'!C[7], ""Open"")"
    
    'To get the longest running openned ticket, filter for openned tickets and return the largest
    If Range("G13").Value = 0 Then
    GoTo Label2
    End If
    Sheets("CORE Network").Select
    Range("A1").AutoFilter Field:=14, Criteria1:="=*Open*", Operator:=xlAnd
    Range("H" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("G14").Select
    ActiveSheet.Paste
    Sheets("CORE Network").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("H14").Select
    ActiveSheet.Paste
    Sheets("CORE Network").Select
    Range("A" & Rows.Count).End(xlUp).Offset(0, 19).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("I14").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("CORE Network").Select
    Selection.AutoFilter
    Sheets("SUMMARY").Select
    
Label2:
    
    Range("G15").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('Core Network'!C[7], ""Closed"")"
    
    'If total tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G12").Value = 0 Then
    Range("G16").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G17").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G18").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H17").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H18").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I17").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I18").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages2
    End If
    
    'If closed tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G15").Value = 0 Then
    Range("G16").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G17").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G18").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H17").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H18").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I17").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I18").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages2
    End If
    
    Range("G16").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(AVERAGE('Core Network'!C[3]))=0,MINUTE(AVERAGE('Core Network'!C[3]))&""mins"",IF(HOUR(AVERAGE('Core Network'!C[3]))>1,HOUR(AVERAGE('Core Network'!C[3]))&""hrs, ""&MINUTE(AVERAGE('Core Network'!C[3]))&""mins"",HOUR(AVERAGE('Core Network'!C[3]))&""hr, ""&MINUTE(AVERAGE('Core Network'!C[3]))&""mins""))"
    Range("G17").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SMALL('Core Network'!C[3],1))=0,MINUTE(SMALL('Core Network'!C[3],1))&""mins"",IF(HOUR(SMALL('Core Network'!C[3],1))>1,HOUR(SMALL('Core Network'!C[3],1))&""hrs, ""&MINUTE(SMALL('Core Network'!C[3],1))&""mins"",HOUR(SMALL('Core Network'!C[3],1))&""hr, ""&MINUTE(SMALL('Core Network'!C[3],1))&""mins""))"
    Range("G18").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(LARGE('Core Network'!C[3],1))=0,MINUTE(LARGE('Core Network'!C[3],1))&""mins"",IF(HOUR(LARGE('Core Network'!C[3],1))>1,HOUR(LARGE('Core Network'!C[3],1))&""hrs, ""&MINUTE(LARGE('Core Network'!C[3],1))&""mins"",HOUR(LARGE('Core Network'!C[3],1))&""hr, ""&MINUTE(LARGE('Core Network'!C[3],1))&""mins""))"
    Range("H17").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Core Network'!C[-7], MATCH(SMALL('Core Network'!C[2],1),'Core Network'!C[2],0))"
    Range("H18").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Core Network'!C[-7], MATCH(LARGE('Core Network'!C[2],1), 'Core Network'!C[2],0))"
    Range("I17").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Core Network'!C[7], MATCH(SMALL('Core Network'!C[1],1), 'Core Network'!C[1],0))"
    Range("I18").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Core Network'!C[7], MATCH(LARGE('Core Network'!C[1],1), 'Core Network'!C[1],0))"
    
SkipAverages2:
    '7c. RAN Network
    Range("G22").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('RAN'!C[-6])-1"
    Range("G23").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('RAN'!C[7], ""Open"")"
    
    'To get the longest running openned ticket, filter for openned tickets and return the largest
    If Range("G23").Value = 0 Then
    GoTo Label3
    End If
    Sheets("RAN").Select
    Range("A1").AutoFilter Field:=14, Criteria1:="=*Open*", Operator:=xlAnd
    Range("H" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("G24").Select
    ActiveSheet.Paste
    Sheets("RAN").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("H24").Select
    ActiveSheet.Paste
    Sheets("RAN").Select
    Range("A" & Rows.Count).End(xlUp).Offset(0, 19).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("I24").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("RAN").Select
    Selection.AutoFilter
    Sheets("SUMMARY").Select
    
Label3:
    Range("G25").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('RAN'!C[7], ""Closed"")"
    
    'If total tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G22").Value = 0 Then
    Range("G26").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G27").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G28").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H27").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H28").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I27").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I28").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages3
    End If
    
    'If closed tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G25").Value = 0 Then
    Range("G26").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G27").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G28").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H27").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H28").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I27").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I28").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages3
    End If
    
    Range("G26").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(AVERAGE('RAN'!C[3]))=0,MINUTE(AVERAGE('RAN'!C[3]))&""mins"",IF(HOUR(AVERAGE('RAN'!C[3]))>1,HOUR(AVERAGE('RAN'!C[3]))&""hrs, ""&MINUTE(AVERAGE('RAN'!C[3]))&""mins"",HOUR(AVERAGE('RAN'!C[3]))&""hr, ""&MINUTE(AVERAGE('RAN'!C[3]))&""mins""))"
    Range("G27").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SMALL('RAN'!C[3],1))=0,MINUTE(SMALL('RAN'!C[3],1))&""mins"",IF(HOUR(SMALL('RAN'!C[3],1))>1,HOUR(SMALL('RAN'!C[3],1))&""hrs, ""&MINUTE(SMALL('RAN'!C[3],1))&""mins"",HOUR(SMALL('RAN'!C[3],1))&""hr, ""&MINUTE(SMALL('RAN'!C[3],1))&""mins""))"
    Range("G28").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(LARGE('RAN'!C[3],1))=0,MINUTE(LARGE('RAN'!C[3],1))&""mins"",IF(HOUR(LARGE('RAN'!C[3],1))>1,HOUR(LARGE('RAN'!C[3],1))&""hrs, ""&MINUTE(LARGE('RAN'!C[3],1))&""mins"",HOUR(LARGE('RAN'!C[3],1))&""hr, ""&MINUTE(LARGE('RAN'!C[3],1))&""mins""))"
    Range("H27").Select
    ActiveCell.FormulaR1C1 = "=INDEX('RAN'!C[-7], MATCH(SMALL('RAN'!C[2],1),'RAN'!C[2],0))"
    Range("H28").Select
    ActiveCell.FormulaR1C1 = "=INDEX('RAN'!C[-7], MATCH(LARGE('RAN'!C[2],1), 'RAN'!C[2],0))"
    Range("I27").Select
    ActiveCell.FormulaR1C1 = "=INDEX('RAN'!C[7], MATCH(SMALL('RAN'!C[1],1), 'RAN'!C[1],0))"
    Range("I28").Select
    ActiveCell.FormulaR1C1 = "=INDEX('RAN'!C[7], MATCH(LARGE('RAN'!C[1],1), 'RAN'!C[1],0))"


SkipAverages3:
    '7d. TX
    Range("G32").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('TX'!C[-6])-1"
    Range("G33").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('TX'!C[7], ""Open"")"
    
    'To get the longest running openned ticket, filter for openned tickets and return the largest
    If Range("G33").Value = 0 Then
    GoTo Label4
    End If
    Sheets("TX").Select
    Range("A1").AutoFilter
    Range("A1").AutoFilter Field:=14, Criteria1:="=*Open*", Operator:=xlAnd
    Range("H" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("G34").Select
    ActiveSheet.Paste
    Sheets("TX").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("H34").Select
    ActiveSheet.Paste
    Sheets("TX").Select
    Range("A" & Rows.Count).End(xlUp).Offset(0, 19).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("I34").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("TX").Select
    Selection.AutoFilter
    Sheets("SUMMARY").Select
    
Label4:
    
    Range("G35").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('TX'!C[7], ""Closed"")"
    
    'If total tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G32").Value = 0 Then
    Range("G36").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G37").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G38").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H37").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H38").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I37").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I38").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages4
    End If
    
    'If total tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G32").Value = 0 Then
    Range("G35").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G37").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G38").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H37").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H38").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I37").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I38").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages4
    End If
    
    Range("G36").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(AVERAGE('TX'!C[3]))=0,MINUTE(AVERAGE('TX'!C[3]))&""mins"",IF(HOUR(AVERAGE('TX'!C[3]))>1,HOUR(AVERAGE('TX'!C[3]))&""hrs, ""&MINUTE(AVERAGE('TX'!C[3]))&""mins"",HOUR(AVERAGE('TX'!C[3]))&""hr, ""&MINUTE(AVERAGE('TX'!C[3]))&""mins""))"
    Range("G37").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SMALL('TX'!C[3],1))=0,MINUTE(SMALL('TX'!C[3],1))&""mins"",IF(HOUR(SMALL('TX'!C[3],1))>1,HOUR(SMALL('TX'!C[3],1))&""hrs, ""&MINUTE(SMALL('TX'!C[3],1))&""mins"",HOUR(SMALL('TX'!C[3],1))&""hr, ""&MINUTE(SMALL('TX'!C[3],1))&""mins""))"
    Range("G38").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(LARGE('TX'!C[3],1))=0,MINUTE(LARGE('TX'!C[3],1))&""mins"",IF(HOUR(LARGE('TX'!C[3],1))>1,HOUR(LARGE('TX'!C[3],1))&""hrs, ""&MINUTE(LARGE('TX'!C[3],1))&""mins"",HOUR(LARGE('TX'!C[3],1))&""hr, ""&MINUTE(LARGE('TX'!C[3],1))&""mins""))"
    Range("H37").Select
    ActiveCell.FormulaR1C1 = "=INDEX('TX'!C[-7], MATCH(SMALL('TX'!C[2],1),'TX'!C[2],0))"
    Range("H38").Select
    ActiveCell.FormulaR1C1 = "=INDEX('TX'!C[-7], MATCH(LARGE('TX'!C[2],1), 'TX'!C[2],0))"
    Range("I37").Select
    ActiveCell.FormulaR1C1 = "=INDEX('TX'!C[7], MATCH(SMALL('TX'!C[1],1), 'TX'!C[1],0))"
    Range("I38").Select
    ActiveCell.FormulaR1C1 = "=INDEX('TX'!C[7], MATCH(LARGE('TX'!C[1],1), 'TX'!C[1],0))"
    Range("I39").Select

SkipAverages4:
    '7e. Telecom
    Range("G42").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('Telecom'!C[-6])-1"
    Range("G43").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('Telecom'!C[7], ""Open"")"
    
    'To get the longest running openned ticket, filter for openned tickets and return the largest
    If Range("G43").Value = 0 Then
    GoTo Label5
    End If
    Sheets("Telecom").Select
    Range("A1").AutoFilter Field:=14, Criteria1:="=*Open*", Operator:=xlAnd
    Range("H" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("G44").Select
    ActiveSheet.Paste
    Sheets("Telecom").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("H44").Select
    ActiveSheet.Paste
    Sheets("Telecom").Select
    Range("A" & Rows.Count).End(xlUp).Offset(0, 19).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("I44").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Telecom").Select
    Selection.AutoFilter
    Sheets("SUMMARY").Select
    
Label5:
    
    Range("G45").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('Telecom'!C[7], ""Closed"")"
    
    'If total tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G42").Value = 0 Then
    Range("G46").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G47").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G48").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H47").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H48").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I47").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I48").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages5
    End If
    
    'If closed tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G45").Value = 0 Then
    Range("G46").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G47").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G48").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H47").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H48").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I47").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I48").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages5
    End If
    
    Range("G46").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(AVERAGE('Telecom'!C[3]))=0,MINUTE(AVERAGE('Telecom'!C[3]))&""mins"",IF(HOUR(AVERAGE('Telecom'!C[3]))>1,HOUR(AVERAGE('Telecom'!C[3]))&""hrs, ""&MINUTE(AVERAGE('Telecom'!C[3]))&""mins"",HOUR(AVERAGE('Telecom'!C[3]))&""hr, ""&MINUTE(AVERAGE('Telecom'!C[3]))&""mins""))"
    Range("G47").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SMALL('Telecom'!C[3],1))=0,MINUTE(SMALL('Telecom'!C[3],1))&""mins"",IF(HOUR(SMALL('Telecom'!C[3],1))>1,HOUR(SMALL('Telecom'!C[3],1))&""hrs, ""&MINUTE(SMALL('Telecom'!C[3],1))&""mins"",HOUR(SMALL('Telecom'!C[3],1))&""hr, ""&MINUTE(SMALL('Telecom'!C[3],1))&""mins""))"
    Range("G48").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(LARGE('Telecom'!C[3],1))=0,MINUTE(LARGE('Telecom'!C[3],1))&""mins"",IF(HOUR(LARGE('Telecom'!C[3],1))>1,HOUR(LARGE('Telecom'!C[3],1))&""hrs, ""&MINUTE(LARGE('Telecom'!C[3],1))&""mins"",HOUR(LARGE('Telecom'!C[3],1))&""hr, ""&MINUTE(LARGE('Telecom'!C[3],1))&""mins""))"
    Range("H47").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Telecom'!C[-7], MATCH(SMALL('Telecom'!C[2],1),'Telecom'!C[2],0))"
    Range("H48").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Telecom'!C[-7], MATCH(LARGE('Telecom'!C[2],1), 'Telecom'!C[2],0))"
    Range("I47").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Telecom'!C[7], MATCH(SMALL('Telecom'!C[1],1), 'Telecom'!C[1],0))"
    Range("I48").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Telecom'!C[7], MATCH(LARGE('Telecom'!C[1],1), 'Telecom'!C[1],0))"
    Range("I49").Select

SkipAverages5:
    '7f. Enterprise
    Range("G52").Select
    Range("G52").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('Enterprise'!C[-6])-1"
    Range("G53").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('Enterprise'!C[7], ""Open"")"
    
    'To get the longest running openned ticket, filter for openned tickets and return the largest
    If Range("G53").Value = 0 Then
    GoTo Label6
    End If
    Sheets("Enterprise").Select
    Range("A1").AutoFilter Field:=14, Criteria1:="=*Open*", Operator:=xlAnd
    Range("H" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("G54").Select
    ActiveSheet.Paste
    Sheets("Enterprise").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("H54").Select
    ActiveSheet.Paste
    Sheets("Enterprise").Select
    Range("A" & Rows.Count).End(xlUp).Offset(0, 19).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("I54").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Enterprise").Select
    Selection.AutoFilter
    Sheets("SUMMARY").Select
    
Label6:
    
    Range("G55").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('Enterprise'!C[7], ""Closed"")"
    
    'If total tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G52").Value = 0 Then
    Range("G56").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G57").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G58").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H57").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H58").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I57").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I58").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages6
    End If
    
    'If closed tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G55").Value = 0 Then
    Range("G56").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G57").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G58").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H57").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H58").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I57").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I58").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages6
    End If
    
    Range("G56").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(AVERAGE('Enterprise'!C[3]))=0,MINUTE(AVERAGE('Enterprise'!C[3]))&""mins"",IF(HOUR(AVERAGE('Enterprise'!C[3]))>1,HOUR(AVERAGE('Enterprise'!C[3]))&""hrs, ""&MINUTE(AVERAGE('Enterprise'!C[3]))&""mins"",HOUR(AVERAGE('Enterprise'!C[3]))&""hr, ""&MINUTE(AVERAGE('Enterprise'!C[3]))&""mins""))"
    Range("G57").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SMALL('Enterprise'!C[3],1))=0,MINUTE(SMALL('Enterprise'!C[3],1))&""mins"",IF(HOUR(SMALL('Enterprise'!C[3],1))>1,HOUR(SMALL('Enterprise'!C[3],1))&""hrs, ""&MINUTE(SMALL('Enterprise'!C[3],1))&""mins"",HOUR(SMALL('Enterprise'!C[3],1))&""hr, ""&MINUTE(SMALL('Enterprise'!C[3],1))&""mins""))"
    Range("G58").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(LARGE('Enterprise'!C[3],1))=0,MINUTE(LARGE('Enterprise'!C[3],1))&""mins"",IF(HOUR(LARGE('Enterprise'!C[3],1))>1,HOUR(LARGE('Enterprise'!C[3],1))&""hrs, ""&MINUTE(LARGE('Enterprise'!C[3],1))&""mins"",HOUR(LARGE('Enterprise'!C[3],1))&""hr, ""&MINUTE(LARGE('Enterprise'!C[3],1))&""mins""))"
    Range("H57").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Enterprise'!C[-7], MATCH(SMALL('Enterprise'!C[2],1),'Enterprise'!C[2],0))"
    Range("H58").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Enterprise'!C[-7], MATCH(LARGE('Enterprise'!C[2],1), 'Enterprise'!C[2],0))"
    Range("I57").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Enterprise'!C[7], MATCH(SMALL('Enterprise'!C[1],1), 'Enterprise'!C[1],0))"
    Range("I58").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Enterprise'!C[7], MATCH(LARGE('Enterprise'!C[1],1), 'Enterprise'!C[1],0))"
    Range("I59").Select

SkipAverages6:
    '7g. Total
    Range("G62").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('Total'!C[-6])-1"
    Range("G63").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('Total'!C[7], ""Open"")"
    
    'To get the longest running openned ticket, filter for openned tickets and return the largest
    If Range("G63").Value = 0 Then
    GoTo Label7
    End If
    Sheets("Total").Select
    Range("A1").AutoFilter Field:=14, Criteria1:="=*Open*", Operator:=xlAnd
    Range("H" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("G64").Select
    ActiveSheet.Paste
    Sheets("Total").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("H64").Select
    ActiveSheet.Paste
    Sheets("Total").Select
    Range("A" & Rows.Count).End(xlUp).Offset(0, 19).Select
    Selection.Copy
    Sheets("SUMMARY").Select
    Range("I64").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Total").Select
    Selection.AutoFilter
    Sheets("SUMMARY").Select
    
Label7:
    
    Range("G65").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('Total'!C[7], ""Closed"")"
    
    'If total tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G62").Value = 0 Then
    Range("G66").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G67").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G68").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H67").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H68").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I67").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I68").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages7
    End If
    
    'If closed tickets = 0 skip the average calculations to avoid dividing by 0 error
    If Range("G65").Value = 0 Then
    Range("G66").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G67").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("G68").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H67").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("H68").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I67").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("I68").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    GoTo SkipAverages7
    End If
    
    Range("G66").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(AVERAGE('Total'!C[3]))=0,MINUTE(AVERAGE('Total'!C[3]))&""mins"",IF(HOUR(AVERAGE('Total'!C[3]))>1,HOUR(AVERAGE('Total'!C[3]))&""hrs, ""&MINUTE(AVERAGE('Total'!C[3]))&""mins"",HOUR(AVERAGE('Total'!C[3]))&""hr, ""&MINUTE(AVERAGE('Total'!C[3]))&""mins""))"
    Range("G67").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SMALL('Total'!C[3],1))=0,MINUTE(SMALL('Total'!C[3],1))&""mins"",IF(HOUR(SMALL('Total'!C[3],1))>1,HOUR(SMALL('Total'!C[3],1))&""hrs, ""&MINUTE(SMALL('Total'!C[3],1))&""mins"",HOUR(SMALL('Total'!C[3],1))&""hr, ""&MINUTE(SMALL('Total'!C[3],1))&""mins""))"
    Range("G68").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(LARGE('Total'!C[3],1))=0,MINUTE(LARGE('Total'!C[3],1))&""mins"",IF(HOUR(LARGE('Total'!C[3],1))>1,HOUR(LARGE('Total'!C[3],1))&""hrs, ""&MINUTE(LARGE('Total'!C[3],1))&""mins"",HOUR(LARGE('Total'!C[3],1))&""hr, ""&MINUTE(LARGE('Total'!C[3],1))&""mins""))"
    Range("H67").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Total'!C[-7], MATCH(SMALL('Total'!C[2],1),'Total'!C[2],0))"
    Range("H68").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Total'!C[-7], MATCH(LARGE('Total'!C[2],1), 'Total'!C[2],0))"
    Range("I67").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Total'!C[7], MATCH(SMALL('Total'!C[1],1), 'Total'!C[1],0))"
    Range("I68").Select
    ActiveCell.FormulaR1C1 = "=INDEX('Total'!C[7], MATCH(LARGE('Total'!C[1],1), 'Total'!C[1],0))"
    Range("I69").Select

SkipAverages7:
    'SUMMARY - THIRD TABLES
    '7h. TOTAL TTs
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('MPLS-IN-VAS'!C[3], ""TELECOM"")"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('MPLS-IN-VAS'!C[2],""CORPORATE"")"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('CORE Network'!C[3], ""TELECOM"")"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('CORE Network'!C[2], ""CORPORATE"")"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('TX'!C[3], ""TELECOM"")"
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('TX'!C[2],""CORPORATE"")"
    Range("L5").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('RAN'!C[3], ""TELECOM"")"
    Range("M5").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('RAN'!C[2], ""CORPORATE"")"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    Range("N4").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    Range("N5").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    Range("L6").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("M6").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("N6").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"

    '7i. TOTAL TT PRIORITIES
    Range("L10").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('MPLS-IN-VAS'!C[-1], ""P1"")"
    Range("M10").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('MPLS-IN-VAS'!C[-2], ""P2"")"
    Range("N10").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('MPLS-IN-VAS'!C[-3], ""P3"")"
    Range("L11").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('CORE Network'!C[-1], ""P1"")"
    Range("M11").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('CORE Network'!C[-2], ""P2"")"
    Range("N11").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('CORE Network'!C[-3], ""P3"")"
    Range("L12").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('TX'!C[-1], ""P1"")"
    Range("M12").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('TX'!C[-2], ""P2"")"
    Range("N12").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('TX'!C[-3], ""P3"")"
    Range("L13").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('RAN'!C[-1], ""P1"")"
    Range("M13").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('RAN'!C[-2], ""P2"")"
    Range("N13").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF('RAN'!C[-3], ""P3"")"
    Range("O10").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("O11").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("O12").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("O13").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("L14").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("M14").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("N14").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("O14").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"

    '7j. CLOSED TTs
    Range("L18").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('MPLS-IN-VAS'!C[-1],""P1"", 'MPLS-IN-VAS'!C[2], ""Closed"")"
    Range("M18").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('MPLS-IN-VAS'!C[-2],""P2"", 'MPLS-IN-VAS'!C[1], ""Closed"")"
    Range("N18").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('MPLS-IN-VAS'!C[-3],""P3"", 'MPLS-IN-VAS'!C, ""Closed"")"
    Range("L19").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('CORE Network'!C[-1],""P1"", 'CORE Network'!C[2], ""Closed"")"
    Range("M19").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('CORE Network'!C[-2],""P2"", 'CORE Network'!C[1], ""Closed"")"
    Range("N19").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('CORE Network'!C[-3],""P3"", 'CORE Network'!C, ""Closed"")"
    Range("L20").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('TX'!C[-1],""P1"", 'TX'!C[2], ""Closed"")"
    Range("M20").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('TX'!C[-2],""P2"", 'TX'!C[1], ""Closed"")"
    Range("N20").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('TX'!C[-3],""P3"", 'TX'!C, ""Closed"")"
    Range("L21").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('RAN'!C[-1],""P1"", 'RAN'!C[2], ""Closed"")"
    Range("M21").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('RAN'!C[-2],""P2"", 'RAN'!C[1], ""Closed"")"
    Range("N21").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('RAN'!C[-3],""P3"", 'RAN'!C, ""Closed"")"
    Range("O18").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("O19").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("O20").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("O21").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("L22").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("M22").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("N22").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("O22").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"

    '7k. OPEN TTs
    Range("L26").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('MPLS-IN-VAS'!C[-1],""P1"", 'MPLS-IN-VAS'!C[2], ""Open"")"
    Range("M26").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('MPLS-IN-VAS'!C[-2],""P2"", 'MPLS-IN-VAS'!C[1], ""Open"")"
    Range("N26").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('MPLS-IN-VAS'!C[-3],""P3"", 'MPLS-IN-VAS'!C, ""Open"")"
    Range("L27").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('CORE Network'!C[-1],""P1"", 'CORE Network'!C[2], ""Open"")"
    Range("M27").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('CORE Network'!C[-2],""P2"", 'CORE Network'!C[1], ""Open"")"
    Range("N27").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('CORE Network'!C[-3],""P3"", 'CORE Network'!C, ""Open"")"
    Range("L28").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('TX'!C[-1],""P1"", 'TX'!C[2], ""Open"")"
    Range("M28").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('TX'!C[-2],""P2"", 'TX'!C[1], ""Open"")"
    Range("N28").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('TX'!C[-3],""P3"", 'TX'!C, ""Open"")"
    Range("L29").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('RAN'!C[-1],""P1"", 'RAN'!C[2], ""Open"")"
    Range("M29").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('RAN'!C[-2],""P2"", 'RAN'!C[1], ""Open"")"
    Range("N29").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS('RAN'!C[-3],""P3"", 'RAN'!C, ""Open"")"
    Range("O26").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("O27").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("O28").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("O29").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("L30").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("M30").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("N30").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("O30").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"


    Sheets("MPLS-IN-VAS").Select
    Selection.AutoFilter
    Range("A1").Select
    Sheets("CORE Network").Select
    Selection.AutoFilter
    Range("A1").Select
    Sheets("TX").Select
    Range("A1").Select
    Sheets("RAN").Select
    Range("A1").Select
    Sheets("Enterprise").Select
    Range("A1").Select
    Sheets("SUMMARY").Select
    Range("A1").Select
    
    'Delete TEMP and Sheet1
    Application.DisplayAlerts = False
    Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
    'Get Tickets Pending from previous day
    Dim ShtIndex As Integer
    
    Dim lcount  As Long
    Dim Rx, Cx As Long

    ShtIndex = 2
    Rx = 3
    Cx = 3
    
    Sheets("RAN").Select
    Sheets("RAN").Move Before:=Sheets(4)
    
    While (ShtIndex <= 5)
    Sheets(ShtIndex).Select
    Range("A1").Select
    lcount = 0
    
    With ActiveSheet
        Set rnData = .UsedRange
        With rnData
            .AutoFilter Field:=14, Criteria1:="=*Open*"
    '        .AutoFilter Field:=19, Criteria1:=">=1"
            .Select
            For Each rngarea In .SpecialCells(xlCellTypeVisible).Areas
                lcount = lcount + rngarea.Rows.Count
            Next
        Sheets("SUMMARY").Cells(Rx, Cx).FormulaR1C1 = lcount - 1
        End With
    End With
    Debug.Print Sheets("SUMMARY").Cells(Rx, Cx).Value
    Selection.AutoFilter
    Range("A1").Select
    
    ShtIndex = ShtIndex + 1
    Rx = Rx + 10
    Wend
    
    Sheets("RAN").Select
    Sheets("RAN").Move Before:=Sheets(6)
    
    Sheets("SUMMARY").Select
    
    'Fill first tables
    'MPLS/IN/VAS
    
    Range("C2").Select
    ActiveCell.FormulaR1C1 = 0
    Range("C4").Select
    ActiveCell.FormulaR1C1 = Range("G5").Value
    Range("C5").Select
    ActiveCell.FormulaR1C1 = Range("G3").Value
    
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SUM('MPLS-IN-VAS'!C[6]))=0,MINUTE(SUM('MPLS-IN-VAS'!C[6]))&""mins"",IF(HOUR(SUM('MPLS-IN-VAS'!C[6]))>1,HOUR(SUM('MPLS-IN-VAS'!C[6]))&""hrs, ""&MINUTE(SUM('MPLS-IN-VAS'!C[6]))&""mins"",HOUR(SUM('MPLS-IN-VAS'!C[6]))&""hr, ""&MINUTE(SUM('MPLS-IN-VAS'!C[6]))&""mins""))"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    'Core Network
    Range("C12").Select
    ActiveCell.FormulaR1C1 = 0
    Range("C14").Select
    ActiveCell.FormulaR1C1 = Range("G15").Value
    Range("C15").Select
    ActiveCell.FormulaR1C1 = Range("G13").Value
    
    Range("D12").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("D13").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("D14").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SUM('Core Network'!C[6]))=0,MINUTE(SUM('Core Network'!C[6]))&""mins"",IF(HOUR(SUM('Core Network'!C[6]))>1,HOUR(SUM('Core Network'!C[6]))&""hrs, ""&MINUTE(SUM('Core Network'!C[6]))&""mins"",HOUR(SUM('Core Network'!C[6]))&""hr, ""&MINUTE(SUM('Core Network'!C[6]))&""mins""))"
    Range("D15").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    'RAN Network
    Range("C22").Select
    ActiveCell.FormulaR1C1 = 0
    Range("C24").Select
    ActiveCell.FormulaR1C1 = Range("G25").Value
    Range("C25").Select
    ActiveCell.FormulaR1C1 = Range("G23").Value
    
    Range("D22").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("D23").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("D24").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SUM('RAN'!C[6]))=0,MINUTE(SUM('RAN'!C[6]))&""mins"",IF(HOUR(SUM('RAN'!C[6]))>1,HOUR(SUM('RAN'!C[6]))&""hrs, ""&MINUTE(SUM('RAN'!C[6]))&""mins"",HOUR(SUM('RAN'!C[6]))&""hr, ""&MINUTE(SUM('RAN'!C[6]))&""mins""))"
    Range("D25").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    'TX Network
    Range("C32").Select
    ActiveCell.FormulaR1C1 = 0
    Range("C34").Select
    ActiveCell.FormulaR1C1 = Range("G35").Value
    Range("C35").Select
    ActiveCell.FormulaR1C1 = Range("G33").Value
    
    Range("D32").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("D33").Select
    ActiveCell.FormulaR1C1 = "N/A"
    Range("D34").Select
    ActiveCell.FormulaR1C1 = "=IF(HOUR(SUM('TX'!C[6]))=0,MINUTE(SUM('TX'!C[6]))&""mins"",IF(HOUR(SUM('TX'!C[6]))>1,HOUR(SUM('TX'!C[6]))&""hrs, ""&MINUTE(SUM('TX'!C[6]))&""mins"",HOUR(SUM('TX'!C[6]))&""hr, ""&MINUTE(SUM('TX'!C[6]))&""mins""))"
    Range("D35").Select
    ActiveCell.FormulaR1C1 = "N/A"
    
    
    'Remove Column S
    ShtIndex = 2
    
    While (ShtIndex <= 9)
    Sheets(ShtIndex).Select
    Columns("S:S").Select
    Selection.Delete Shift:=xlLeft
    Selection.Delete Shift:=xlLeft
    Range("A1").Select
    
    ShtIndex = ShtIndex + 1
    Wend
    
    'Remove extra title column from Telecom and Total if there exist
    Sheets("Telecom").Select
    Range("A1").Select
    If Sheets("Telecom").Range("A2").Value = "Ticket ID" Then
    Sheets("Telecom").Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    End If
    
    Sheets("Total").Select
    Range("A1").Select
    If Sheets("Total").Range("A2").Value = "Ticket ID" Then
    Sheets("Total").Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    End If


    'FORMAT SUMMARY SHEET
    
    Sheets("SUMMARY").Select
    Range("A1:O68").Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Columns("C:C").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("G:G").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("D:D").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("H:H").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("B:B").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Range("A1").Select
    
    
    
    
    Application.ScreenUpdating = True

End Sub
