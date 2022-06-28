VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
On Error Resume Next

Application.Interactive = False
Application.ScreenUpdating = False

Workbooks("theunitCircle").Worksheets("circle").Select

DoEvents

CommandBars.ExecuteMso "MinimizeRibbon"

    With Application
        .DisplayFormulaBar = False
        .ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",false)"
        .CommandBars("Worksheet Menu Bar").Enabled = False
        .DisplayStatusBar = False
        .DisplayScrollBars = False
        .Width = 906
        .Height = 598
    End With
    
    DoEvents
    
    With Windows("theunitCircle.xlsm")
        .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .Zoom = 50
    End With

DoEvents
Application.Interactive = True
Application.ScreenUpdating = True

End Sub


