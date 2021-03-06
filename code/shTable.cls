VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public FoundWordsCount As Integer               'Kiek �od�i� jau i�braukta
Public UsedHintCount As Integer

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim c As Range
    Dim i As Integer
    Dim strSlctd As String
    Dim strSlctdRvrs As String

    ' In case sheet is unprotected, don't let code affect word list
    If Not Intersect(Target, Range("rngTableWords")) Is Nothing Then Exit Sub
    If DEVELOPMODE Then Exit Sub
    ActiveSheet.Unprotect
    '�ra�o pasiriktos srities raides � strSlctd
    For Each c In Target.Cells
        strSlctd = strSlctd & c.Value
    Next c

    'U�ra�om pa�ym�t� srit� strSlctd atbulai
    For i = Len(strSlctd) To 1 Step -1
        strSlctdRvrs = strSlctdRvrs & Mid(strSlctd, i, 1)
    Next i

    'Tikrinam kiekvien� �od� ar sutampa
    For Each c In Range("rngTableWords").Cells
        With c
            If .Value = strSlctd Or .Value = strSlctdRvrs Then
                .Font.Strikethrough = True
                .Interior.ColorIndex = 34               'Nuspalvinam �od� �onin�j �od�i� s�ra�e
                Target.Interior.ColorIndex = 34         'Nuspalvinam �od� lentel�j
                FoundWordsCount = FoundWordsCount + 1   'Sekam kiek surasta �od�i�

                If FoundWordsCount = 20 Then
                    Call modGame.ClearTable
                    shResults.Activate
                    Call modResults.Results
                    FoundWordsCount = 0
                    UsedHintCount = 0
                    Exit Sub
                End If
            End If
        End With
    Next c

    '�ra�om �od�ius i� masyvo � lentel�
    Dim j As Integer
    For j = LBound(arrTable, 2) To UBound(arrTable, 2)
        For i = LBound(arrTable, 1) To UBound(arrTable, 1)
            Cells(j + 1, i + 1) = arrTable(j, i)
        Next i
    Next j

    'Per nauj� �ra�om �od�ius � pagalbin� �od�i� lentel�
    For i = LBound(arrWords, 1) + 1 To UBound(arrWords, 1)
        Cells(i, HelpWordsColumn) = arrWords(i, WordListID)
    Next i

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Function getUsedHintCount() As Integer
    getUsedHintCount = UsedHintCount
End Function

Private Sub CheckBox1_Click()
    Call modAnswers.ShowAnswers
End Sub
