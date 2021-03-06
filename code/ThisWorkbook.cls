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
' Programavim� pagrind� projektinis darbas
' �mogaus - Kompiuterio s�sajos projektinis darbas
'
' autoriai:
' Alius Sultanovas
' Paulius �riub�nas
' Lukas Vitkevi�ius
'
' 2014-05-20
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub Workbook_Activate()
    Application.ScreenUpdating = False                      'i�jungia Macro atvaizdavim�
    Application.DisplayAlerts = False
    shOptions.Activate                          'atveria shOptions lap�
    With ActiveWindow
        .DisplayHorizontalScrollBar = False     'paslepia Horizontal� scrollbar
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
    Application.DisplayFormulaBar = True    '�jungia atgal
    Application.ScreenUpdating = True       '�jungia atgal
    Application.DisplayStatusBar = True
    Application.RollZoom = False
    Application.ActiveWindow.Zoom = 100     'Atstato zoom� � 100%
    shOptions.Activate
    ThisWorkbook.Save
    ThisWorkbook.Saved = True       'I�saugo Workbook
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Excel.Window)
    Wn.Caption = "�od�i� i�braukymas!"
End Sub
