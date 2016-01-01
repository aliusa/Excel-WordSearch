VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Vilniaus Kooperacijos Kolegija
' ISDP-13
' Programavimø pagrindø projektinis darbas
' Þmogaus - Kompiuterio sàsajos projektinis darbas
'
' autoriai:
' Alius Sultanovas
' Paulius Ðriubënas
' Lukas Vitkevièius
'
' 2014-05-20
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub Workbook_Activate()
    Application.ScreenUpdating = False                      'iðjungia Macro atvaizdavimà
    Application.DisplayAlerts = False
    shOptions.Activate                          'atveria shOptions lapà
    With ActiveWindow
        .DisplayHorizontalScrollBar = False     'paslepia Horizontalø scrollbar
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = False
        .DisplayHeadings = False
        .DisplayGridlines = False
    End With
    Application.DisplayFormulaBar = False
    Application.RollZoom = True                 'Scroll - zoomina
    Application.DisplayStatusBar = False
    Application.DisplayFullScreen = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Call modGame.ClearTable
    Application.DisplayFormulaBar = True    'Ájungia atgal
    Application.ScreenUpdating = True       'Ájungia atgal
    Application.DisplayStatusBar = True
    Application.RollZoom = False
    Application.ActiveWindow.Zoom = 100     'Atstato zoomà á 100%
    shOptions.Activate
    ThisWorkbook.Save
    ThisWorkbook.Saved = True       'Iðsaugo Workbook
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Excel.Window)
    Wn.Caption = "Þodþiø iðbraukymas!"
End Sub
