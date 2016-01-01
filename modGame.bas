Attribute VB_Name = "modGame"
Option Explicit                         'Priverèia deklaruoti kintamuosius

'Public kintamieji, kad juos matytø visame modulyje
Public TableSize As Integer             'Lentelës dydis
Public WordListID As Integer            'Þodþiø sàraðas
Public FontCase As Integer              'Ðrifto dydis

Public arrWords() As Variant            'Visi þodþiai su pavadinimais
Public arrTable() As String             'Lentelë

Public WordLength As Integer            'Einamojo þodþio ilgis
Public Fits As Integer                  'Loginis kintamasis á kurià pusæ paraðytas þodis
Public DEVELOPMODE As Boolean           'DEBUG

Public WordDirection As Integer         'Kryptis á kurià eina þodis
Public arrWordsX As Integer             'Kurioj eilutëj prasideda þodis
Public arrWordsY As Integer             'Kuriam stulpely prasideda þodis
Public HelpWordsColumn As Integer       'Stulepelis á kurá áraðo þodþius

Dim rndX As Integer
Dim rndY As Integer

Public sngStart As Single

Function StartGame()
    shTable.Activate                                        'atidarom shTable lapà
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayGridlines = False
        .DisplayWorkbookTabs = False                         'DEFAULT - FALSE
        .DisplayVerticalScrollBar = False
        .DisplayHorizontalScrollBar = False
    End With
    
    DEVELOPMODE = shOptions.Cells(1, 4).Value
    If DEVELOPMODE Then ActiveWindow.DisplayWorkbookTabs = True
    
    shWords.Visible = xlSheetVisible
    shOptions.Visible = xlSheetVisible
    shTable.Visible = xlSheetVisible
    shResults.Visible = xlSheetVisible
    
    'Nustatom lentelës dydá
    TableSize = Val(shOptions.Range("rngOptionsSize").Value)
    
    'Kuri þodþiø kategorija buvo pasirinkta
    WordListID = shWords.Range("rngWordsListID")
    
    'Didþiosios ar maþiosios raidës bus lentelëj
    FontCase = IIf(shWords.Range("rngWordsCase").Value = 1, 1, 2)   'True - Didþ, False - maþ
    
    Call modGame.ClearTable                             'iðvalom lentelæ
    
    If TableSize <> 30 Then
        shTable.Range(Cells(TableSize + 2, TableSize + 2), Cells(31, 31)).EntireRow.Hidden = True
        shTable.Range(Cells(TableSize + 2, TableSize + 2), Cells(31, 31)).EntireColumn.Hidden = True
    End If
    
    WordDirection = 23  'Stulpelis á kurá bus áraðyta þodþiø krypties vieta
    arrWordsX = 24      'Stulpelis á kurá bus áraðyta þodþiø eilutës vieta
    arrWordsY = 25      'Stulpelis á kurá bus áraðyta þodþiø stulpelio vieta
    HelpWordsColumn = 33 'Stulpelis kuriame bus áraðyti pagalbiniai þodþiai
    
    Dim strChr As String
    strChr = "AÀBCÈDEÆËFGHIÁYJKLMNOPRSÐTUØÛVZÞaàbcèdeæëfghiáyjklmnoprsðtuøûvzþ"
    Dim arrChars(1 To 32, 1 To 2) As String             'Dvimatis raidþiø masyvas
    Call modGame.PutChrsToArr(strChr, arrChars)         'Ádedam raides á dvimatá masyvà
    
    arrWords() = shWords.Range("listWords")             'Ákeliam visus þodþius á Variant masyvà
    
    If FontCase = 2 Then        'Jei maþosios - pakeisti þodþiø raides maþosiom
        Dim i12 As Integer
        For i12 = LBound(arrWords(), 1) + 1 To UBound(arrWords(), 1) '+1 - pradedam nuo 1-o þodþio
            arrWords(i12, WordListID) = LCase(arrWords(i12, WordListID))
        Next i12
    End If
    
    ReDim arrTable(1 To TableSize, 1 To TableSize) As String   'Lentelë pagal ávestà lentelës dydá
    
    Dim i3 As Integer
    Dim arrWordsList() As String
    For i3 = LBound(arrWords, 1) + 1 To UBound(arrWords, 1) 'áraðo þodþius á ðoninæ lente
        If FontCase = 1 Then
            shTable.Cells(i3, HelpWordsColumn) = arrWords(i3, WordListID) 'didþiosios
            Else: shTable.Cells(i3, HelpWordsColumn) = LCase(arrWords(i3, WordListID)) 'maþosios
        End If
    Next i3
    
    Dim word As String                                      'einamasis þodis
    Dim i4 As Integer
    For i4 = LBound(arrWords, 1) + 1 To UBound(arrWords, 1)
        rndX = RndNum(TableSize)                            'randominam naujà X
        rndY = RndNum(TableSize)                            'randominam naujà Y
        
        Dim i5 As Integer
        Do Until IsEmpty(arrTable(rndX, rndY)) Or (i5 = 50) 'Jei elementas uþimtas - generuojam naujà vietà
            i5 = i5 + 1
            rndX = RndNum(TableSize)                        'randominam naujà X
            rndY = RndNum(TableSize)                        'randominam naujà Y
        Loop
        
        WordLength = Len(arrWords(i4, WordListID))          'Þodþio ilgis
        
        Fits = IsItFitIn(WordLength)                        'Tikrinam ar telpa á lentelæ
        Fill Fits, WordLength, i4                     'Jei telpa - bandom áraðyt arba kartojam
    Next i4
    
    If Not DEVELOPMODE Then Call ChrsToArrTable(arrChars()) 'Áraðom á likusias vietas atsitiktines raides
    
    shTable.Protect                                         'Uþrakinam shTable lapà
    
    shResults.Unprotect
    shResults.Cells(1, 7) = Timer()
    shResults.Cells(1, 9) = Now()
    shResults.Protect
End Function

Function IsItFitIn(ByRef wordLen As Integer)
    Dim direction As Integer
    direction = RndNum(4)
    Select Case direction
        Case 1                                          'UP
            If (rndX - wordLen) >= 1 Then
                IsItFitIn = 1                           'Telpa á virðø
                Else: IsItFitIn = -1                    'Netelpa á virðø
            End If
        Case 2                                          'RIGHT
            If (rndY + wordLen) <= TableSize + 1 Then
                IsItFitIn = 2                           'Telpa á deðinæ
                Else: IsItFitIn = -2                    'Netelpa á deðinæ
            End If
        Case 3                                          'DOWN
            If (rndX + wordLen) <= TableSize + 1 Then
                IsItFitIn = 3                           'Telpa á apaèià
                Else: IsItFitIn = -3                    'Netelpa á apaèià
            End If
        Case 4                                          'LEFT
            If (rndY - wordLen) >= 1 Then
                IsItFitIn = 4                           'Telpa á kairæ
                Else: IsItFitIn = -4                    'Netelpa á kairæ
            End If
        Case Else
            MsgBox "ERROR FitsIn Select Case Else"
    End Select
End Function

Function Fill(ByRef Fits As Integer, ByRef WordLength As Integer, ByRef i4 As Integer)
    Select Case Fits
        Case 1                                                  'FILL UP
            Dim uzimta1 As Boolean
            Dim i6 As Integer
            For i6 = 0 To WordLength - 1
                If Not IsEmpty(Cells(rndX + 1 - i6, rndY + 1)) Then
                    Call DoItAgainBoy(i4)                       'Darom rekursijà
                    uzimta1 = True                              'Þodis netelpa á virðø
                    Exit For
                End If
            Next i6
            If uzimta1 = False Then                             'Áraðom þodá á virðø
                For i6 = 0 To WordLength - 1
                    Cells(rndX + 1 - i6, rndY + 1) = Mid(arrWords(i4, WordListID), i6 + 1, 1)
                    Call WriteCoordsToArray(i4)
                Next i6
            End If
        Case 2                                                  'FILL RIGHT
            Dim uzimta2 As Boolean
            Dim i7 As Integer
            For i7 = 0 To WordLength - 1
                If Not IsEmpty(Cells(rndX + 1, rndY + 1 + i7)) Then
                    Call DoItAgainBoy(i4)                       'Darom rekursijà
                    uzimta2 = True                              'Þodis netelpa á deðinæ
                    Exit For
                End If
            Next i7
            If uzimta2 = False Then                             'Áraðom þodá á deðinæ
                For i7 = 0 To WordLength - 1
                    Cells(rndX + 1, rndY + 1 + i7) = Mid(arrWords(i4, WordListID), i7 + 1, 1)
                    Call WriteCoordsToArray(i4)
                Next i7
            End If
        Case 3                                                  'FILL DOWN
            Dim uzimta3 As Boolean
            Dim i8 As Integer
            For i8 = 0 To WordLength - 1
                If Not IsEmpty(Cells(rndX + 1 + i8, rndY + 1)) Then
                    Call DoItAgainBoy(i4)                       'Darom rekursijà
                    uzimta3 = True                              'Þodis netelpa á apaèià
                    Exit For
                End If
            Next i8
            If uzimta3 = False Then                             'Áraðom þodá á apaèià
                For i8 = 0 To WordLength - 1
                    Cells(rndX + 1 + i8, rndY + 1) = Mid(arrWords(i4, WordListID), i8 + 1, 1)
                    Call WriteCoordsToArray(i4)
                Next i8
            End If
        Case 4                                                  'FILL LEFT
            Dim uzimta4 As Boolean
            Dim i9 As Integer
            For i9 = 0 To WordLength - 1
                If Not IsEmpty(Cells(rndX + 1, rndY + 1 - i9)) Then
                    Call DoItAgainBoy(i4)                       'Darom rekursijà
                    uzimta4 = True                              'Þodis nelepa á kairæ
                    Exit For
                End If
            Next i9
            If uzimta4 = False Then                             'Áraðom þodá á kairæ
                For i9 = 0 To WordLength - 1
                    Cells(rndX + 1, rndY + 1 - i9) = Mid(arrWords(i4, WordListID), i9 + 1, 1)
                    Call WriteCoordsToArray(i4)
                Next i9
            End If
        Case -1, -2, -3, -4                                     'Netelpa - ið naujo ieðkom vietos
            Dim i10 As Integer
            i10 = 1
            Do Until (Fits > 0) Or (i10 = 10)
                rndX = RndNum(TableSize)
                rndY = RndNum(TableSize)
                Fits = IsItFitIn(WordLength)
                i10 = i10 + 1
            Loop
            Fill Fits, WordLength, i4
        Case Else
            MsgBox "ERROR FitItIn Select Case Else"
    End Select
End Function

Function PutChrsToArr(ByRef strChar As String, ByRef strArr() As String) As String()
    Dim i As Integer
    For i = 1 To 32                             'iteruojam strChar
        strArr(i, 1) = Mid(strChar, i, 1)       '1 masyvo dimensija - didþiosios raidës
    Next i
    For i = 33 To 64                            'iteruojam strChar
        strArr(i - 32, 2) = Mid(strChar, i, 1)  '2 masyvo dimensija - maþosios raidës
    Next i
    PutChrsToArr = strArr()                     'gràþinam masyvà
End Function

Function ChrsToArrTable(ByRef arrChars() As String)
    'Ádedam random raides á likusias masyvo ir lentelës vietas
    Dim lrow As Integer
    Dim lcol As Integer
    Dim rndChar As Integer
    For lcol = LBound(arrTable, 2) To UBound(arrTable, 2)           'dvimatis masyvas
        For lrow = LBound(arrTable, 1) To UBound(arrTable, 1)       'vienmatis masyvas
            If (Cells(lcol + 1, lrow + 1) = "") Then                'jei celë tuðèia
                rndChar = RndNum(32)                                'generuojam raidæ
                arrTable(lcol, lrow) = arrChars(rndChar, FontCase)  'áraðom raidæ DIDÞ/maþ
                Cells(lcol + 1, lrow + 1) = arrTable(lcol, lrow)    'áraðom masyvo elementà á lentelæ
                Else: arrTable(lcol, lrow) = Cells(lcol + 1, lrow + 1)   'áraðom celæ á masyvà
            End If
        Next lrow
    Next lcol
End Function

Function ClearTable()
    shWords.Range("rngWordsCheckBox1").Value = False    'Nuimà varnelæ nuo CheckBox
    shTable.Unprotect                                   'Nuimam apsaugà
    shTable.Range("rngTableGrid").ClearContents
    shTable.Range("rngTableGrid").Interior.Color = xlNone
    shTable.Range("rngTableWords").ClearContents
    shTable.Range("rngTableWords").Interior.Color = xlNone
    shTable.Range("rngTableWords").Font.Strikethrough = False
    shTable.Range(Cells(2, 2), Cells(32, 32)).EntireRow.Hidden = False
    shTable.Range(Cells(2, 2), Cells(32, 32)).EntireColumn.Hidden = False
End Function

Function RndNum(ByRef n As Integer) As Integer
    Randomize                   'randominam skaièiø
    RndNum = Int(n * Rnd) + 1   'iðvedam randomintà skaièiø
End Function

Function HintMe()
    If DEVELOPMODE Then Exit Function                   'DEBUG
    If ActiveCell.Interior.Color <> 16777215 Then Exit Function 'Jei langelis spalvotas - iðeiti ið f-jos
    If ActiveCell.Column <> HelpWordsColumn Then Exit Function  'Jei nepaþymëtas þodþiø stulpelis
    
    Dim x As Integer
    Dim y As Integer
    x = arrWords(ActiveCell.Row, arrWordsX)     'Nuskaitom eilutæ kur prasida þodis
    y = arrWords(ActiveCell.Row, arrWordsY)     'Nuskaitom stulpelá kur prasideda þodis
    
    If Cells(x, y).Interior.Color <> 16777215 Then Exit Function
    
    shTable.Unprotect                          'Nuimam apsaugà
    Cells(x, y).Interior.Color = vbYellow      'Nuspalvinam
    shTable.UsedHintCount = shTable.UsedHintCount + 1
    shTable.Protect                             'Uþdedam atgal apsaugà
End Function

Function WriteCoordsToArray(ByRef i4 As Integer)
    arrWords(i4, arrWordsX) = rndX + 1          'Áraðom kurioje eilutëje prasideda þodis
    arrWords(i4, arrWordsY) = rndY + 1          'Áraðom kuriame stulpelyje prasideda þodis
    arrWords(i4, WordDirection) = Fits          'Á kurià kryptá paraðytas þodis
End Function

Function DoItAgainBoy(ByRef i4 As Integer)
    rndX = RndNum(TableSize)        'Generuojam naujà X
    rndY = RndNum(TableSize)        'Generuojam naujà Y
    Fits = IsItFitIn(WordLength)    'Patikrinam ar telpa
    Fill Fits, WordLength, i4       'Rekursuojam
End Function

Function getTableSize() As Integer
    getTableSize = TableSize
End Function

Function getWordsCategory() As String
    getWordsCategory = arrWords(1, WordListID)
End Function

Function getFontSize() As Integer
    getFontSize = FontCase
End Function
