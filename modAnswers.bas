Attribute VB_Name = "modAnswers"
Option Explicit

Function ShowAnswers()
    shTable.Unprotect
    If shWords.Range("rngWordsCheckBox1") = True Then
        shTable.Range("rngTableWords").Interior.Color = xlNone
        shTable.Range("rngTableWords").Font.Strikethrough = False
        
        Dim i As Integer, x As Integer, y As Integer, dir As Integer, j As Integer
        
        For i = 1 To 20
            dir = arrWords(i + 1, modGame.WordDirection)
            x = arrWords(i + 1, modGame.arrWordsX)
            y = arrWords(i + 1, modGame.arrWordsY)
            
            For j = 0 To Len(arrWords(i + 1, WordListID)) - 1
                Select Case dir
                    Case 1: shTable.Cells(x - j, y).Interior.ColorIndex = 8
                    Case 2: shTable.Cells(x, y + j).Interior.ColorIndex = 3
                    Case 3: shTable.Cells(x + j, y).Interior.ColorIndex = 7
                    Case 4: shTable.Cells(x, y - j).Interior.ColorIndex = 6
                End Select
            Next j
        Next i
        
        Else: cmdReset_Click
    End If
    shTable.Protect
End Function
