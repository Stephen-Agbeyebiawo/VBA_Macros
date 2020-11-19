Sub RAN_MS_REPORT()
'
' RAN_MS_REPORT Macro
'

'
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    Sheets("SPU").Select
    Range("A1").Select
    Sheets("SPU").Select
    Range("A2").Select
    Windows("1 RAN SUPPORT REPORT_ BSC6900 UMTS SPU CPU LOAD.xlsx").Activate
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A2").Select
    Sheets("EGPU").Select
    Range("A2").Select
    Windows("2 RAN SUPPORT REPORT_ BSC6910UMTS EGPU CPU LOAD.xlsx").Activate
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A2").Select
    Sheets("DPU").Select
    Range("A2").Select
    Windows("3 RAN SUPPORT REPORT_ BSC6900 UMTS DPU DSP USAGE.xlsx").Activate
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A2").Select
    Sheets("TRANS").Select
    Range("A2").Select
    Windows("4 RAN SUPPORT REPORT_ BSC6900UMTS INTERFACE BOARD TRANSLOAD.xlsx"). _
        Activate
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    Windows("4 RAN SUPPORT REPORT_ BSC6900UMTS INTERFACE BOARD TRANSLOAD.xlsx"). _
        Activate
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    Windows("4 RAN SUPPORT REPORT_ BSC6900UMTS INTERFACE BOARD TRANSLOAD.xlsx"). _
        Activate
    Range("A2:J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A2").Select
    Sheets("EXOU").Select
    Range("A2").Select
    Windows("5 RAN SUPPORT REPORT_ BSC6910 UO EXOU TRANSLOAD.xlsx").Activate
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    Windows("5 RAN SUPPORT REPORT_ BSC6910 UO EXOU TRANSLOAD.xlsx").Activate
    Range("A2:J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    ActiveSheet.Paste
    Range("A2").Select
    Sheets("PORTS IUPS Raw Data").Select
    Range("A2").Select
    Windows("6 RAN_SUPPORT_REPORT_INTERFACE_BOARDS_PORT_UTILIZATION.xlsx").Activate
    Range("A2:F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("6 RAN_SUPPORT_REPORT_INTERFACE_BOARDS_PORT_UTILIZATION.xlsx").Activate
    Selection.End(xlUp).Select
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    Range("G2").Select
    ActiveSheet.Paste
    Windows("6 RAN_SUPPORT_REPORT_INTERFACE_BOARDS_PORT_UTILIZATION.xlsx").Activate
    Selection.End(xlUp).Select
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    Sheets("PORTS IUB Raw Data").Select
    Range("G2").Select
    ActiveSheet.Paste
    Range("A2").Select
    Windows("6 RAN_SUPPORT_REPORT_INTERFACE_BOARDS_PORT_UTILIZATION.xlsx").Activate
    Selection.End(xlUp).Select
    Range("A2:F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    ActiveSheet.Paste
    Range("A2").Select
    Sheets("SCU").Select
    Range("A2").Select
    Windows("7 RAN SUPPORT REPORT_ SCU INTER_SUB RACK UTILIZATION.xlsx").Activate
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("VODAFONE RAN MS REPORT.xlsx").Activate
    ActiveSheet.Paste
    Range("A2").Select
    Sheets("SPU_CHARTS").Select
    ActiveSheet.ChartObjects("Chart 1").Activate
    Application.CutCopyMode = False
    ActiveWorkbook.RefreshAll
    Windows("1 RAN SUPPORT REPORT_ BSC6900 UMTS SPU CPU LOAD.xlsx").Close
    Windows("2 RAN SUPPORT REPORT_ BSC6910UMTS EGPU CPU LOAD.xlsx").Close
    Windows("3 RAN SUPPORT REPORT_ BSC6900 UMTS DPU DSP USAGE.xlsx").Close
    Windows("4 RAN SUPPORT REPORT_ BSC6900UMTS INTERFACE BOARD TRANSLOAD.xlsx").Close
    Windows("5 RAN SUPPORT REPORT_ BSC6910 UO EXOU TRANSLOAD.xlsx").Close
    Windows("6 RAN_SUPPORT_REPORT_INTERFACE_BOARDS_PORT_UTILIZATION.xlsx").Close
    Windows("7 RAN SUPPORT REPORT_ SCU INTER_SUB RACK UTILIZATION.xlsx").Close
    
End Sub
