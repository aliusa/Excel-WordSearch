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
    'Apsauga, jog iðtrynus/pakeitus Lentelës dydþio celës skaièiø - atsatytø á 20
    If (shOptions.Range("rngOptionsSize").Value = xlNone) _
        Or Not IsNumeric(shOptions.Range("rngOptionsSize").Value) _
        Or (shOptions.Range("rngOptionsSize").Value < 20) _
        Or (shOptions.Range("rngOptionsSize").Value > 30) _
        Then
        shOptions.Range("rngOptionsSize").Value = 20
    End If

    'Jei Þodþiø Kategorijos reikðmë tuðèia - áraðo 2 reikðmæ ið listWordsHeaders ið shWords lapo
    If Not IsNumeric(shWords.Range("rngWordsListID").Value) Then
        shOptions.Range("rngOptionsWordList").Value = shWords.Range("listWordsHeaders")(2)
    End If

    'Jei Ðrifto Dydis reikðmë tuðèia - áraðo 1à reikðmæ (DIDÞIOSIOS), ið listWordsCase
    If shOptions.Range("rngOptionsCase").Value = "" Then
        shOptions.Range("rngOptionsCase").Value = shWords.Range("listWordsCase")(1)
    End If
End Sub
