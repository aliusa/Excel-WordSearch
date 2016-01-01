Attribute VB_Name = "modButtons"
Option Explicit

Sub cmdStartGame_Click()
    Call modGame.StartGame
End Sub

Sub cmdReset_Click()
    Call modGame.ClearTable
    shTable.FoundWordsCount = 0
    shOptions.Activate
End Sub

Sub cmdHint_Click()
    Call modGame.HintMe
End Sub

Sub cmdGoBack_Click()
    shOptions.Activate
    Application.RollZoom = True         'Scroll = zoomina
End Sub

Sub cmdResults_Click()
    shResults.Unprotect
    shResults.Cells(1, 7).ClearContents
    shResults.Cells(1, 9).ClearContents
    shResults.Protect
    shResults.Activate
    Application.RollZoom = False
End Sub
