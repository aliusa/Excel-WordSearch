VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Apsauga, jog i�trynus/pakeitus Lentel�s dyd�io cel�s skai�i� - atsatyt� � 20
    If (shOptions.Range("rngOptionsSize").Value = xlNone) _
        Or Not IsNumeric(shOptions.Range("rngOptionsSize").Value) _
        Or (shOptions.Range("rngOptionsSize").Value < 20) _
        Or (shOptions.Range("rngOptionsSize").Value > 30) _
        Then
        shOptions.Range("rngOptionsSize").Value = 20
    End If

    'Jei �od�i� Kategorijos reik�m� tu��ia - �ra�o 2 reik�m� i� listWordsHeaders i� shWords lapo
    If Not IsNumeric(shWords.Range("rngWordsListID").Value) Then
        shOptions.Range("rngOptionsWordList").Value = shWords.Range("listWordsHeaders")(2)
    End If

    'Jei �rifto Dydis reik�m� tu��ia - �ra�o 1� reik�m� (DID�IOSIOS), i� listWordsCase
    If shOptions.Range("rngOptionsCase").Value = "" Then
        shOptions.Range("rngOptionsCase").Value = shWords.Range("listWordsCase")(1)
    End If
End Sub
