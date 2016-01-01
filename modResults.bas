Attribute VB_Name = "modResults"
Option Explicit

Public lrow As Integer

Public Sub Results()
    shResults.Unprotect
    formVardas.Show
    
    Dim Size As Integer
    Size = getTableSize()
    
    Dim Category As String
    Category = getWordsCategory()
    
    Dim Font As Integer
    Font = getFontSize()
    Dim FontCas As String
    FontCas = IIf(Font = 1, "Didþiosios", "maþosios")
    
    Dim sngStart As Single, sngEnd As Single, sngElapsed As Single
    sngStart = shResults.Cells(1, 7).Value
    sngEnd = Timer()
    sngElapsed = Format(sngEnd - sngStart)
    shResults.Cells(1, 7).ClearContents
    Dim dateNow As String
    dateNow = shResults.Cells(1, 9).Value
    shResults.Cells(1, 9).ClearContents
    
    Dim mm As Integer, ss As Integer
    mm = (sngElapsed / 60)
    ss = sngElapsed Mod 60
    
    Dim UsedHintCount As Integer
    UsedHintCount = shTable.getUsedHintCount()
    
    lrow = 3
    Do Until IsEmpty(shResults.Cells(lrow, 2))
        lrow = lrow + 1
    Loop
    
    shResults.Cells(lrow, 2) = lrow - 2
    shResults.Cells(lrow, 4) = Size
    shResults.Cells(lrow, 5) = Category
    shResults.Cells(lrow, 6) = FontCas
    shResults.Cells(lrow, 7) = mm & ":" & ss
    shResults.Cells(lrow, 8) = UsedHintCount
    shResults.Cells(lrow, 9) = dateNow
    shResults.Protect
End Sub
