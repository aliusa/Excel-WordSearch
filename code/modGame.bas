Attribute VB_Name = "modGame"
Option Explicit                         'Priver�ia deklaruoti kintamuosius

'Public kintamieji, kad juos matyt� visame modulyje
Public TableSize As Integer             'Lentel�s dydis
Public WordListID As Integer            '�od�i� s�ra�as
Public FontCase As Integer              '�rifto dydis

Public arrWords() As Variant            'Visi �od�iai su pavadinimais
Public arrTable() As String             'Lentel�

Public WordLength As Integer            'Einamojo �od�io ilgis
Public Fits As Integer                  'Loginis kintamasis � kuri� pus� para�ytas �odis
Public DEVELOPMODE As Boolean           'DEBUG

Public WordDirection As Integer         'Kryptis � kuri� eina �odis
Public arrWordsX As Integer             'Kurioj eilut�j prasideda �odis
Public arrWordsY As Integer             'Kuriam stulpely prasideda �odis
Public HelpWordsColumn As Integer       'Stulepelis � kur� �ra�o �od�ius

Dim rndX As Integer
Dim rndY As Integer

Public sngStart As Single

Function StartGame()
    shTable.Activate                                        'atidarom shTable lap�
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
    
    'Nustatom lentel�s dyd�
    TableSize = Val(shOptions.Range("rngOptionsSize").Value)
    
    'Kuri �od�i� kategorija buvo pasirinkta
    WordListID = shWords.Range("rngWordsListID")
    
    'Did�iosios ar ma�iosios raid�s bus lentel�j
    FontCase = IIf(shWords.Range("rngWordsCase").Value = 1, 1, 2)   'True - Did�, False - ma�
    
    Call modGame.ClearTable                             'i�valom lentel�
    
    If TableSize <> 30 Then
        shTable.Range(Cells(TableSize + 2, TableSize + 2), Cells(31, 31)).EntireRow.Hidden = True
        shTable.Range(Cells(TableSize + 2, TableSize + 2), Cells(31, 31)).EntireColumn.Hidden = True
    End If
    
    WordDirection = 23  'Stulpelis � kur� bus �ra�yta �od�i� krypties vieta
    arrWordsX = 24      'Stulpelis � kur� bus �ra�yta �od�i� eilut�s vieta
    arrWordsY = 25      'Stulpelis � kur� bus �ra�yta �od�i� stulpelio vieta
    HelpWordsColumn = 33 'Stulpelis kuriame bus �ra�yti pagalbiniai �od�iai
    
    Dim strChr As String
    strChr = "A�BC�DE��FGHI�YJKLMNOPRS�TU��VZ�a�bc�de��fghi�yjklmnoprs�tu��vz�"
    Dim arrChars(1 To 32, 1 To 2) As String             'Dvimatis raid�i� masyvas
    Call modGame.PutChrsToArr(strChr, arrChars)         '�dedam raides � dvimat� masyv�
    
    arrWords() = shWords.Range("listWords")             '�keliam visus �od�ius � Variant masyv�
    
    If FontCase = 2 Then        'Jei ma�osios - pakeisti �od�i� raides ma�osiom
        Dim i12 As Integer
        For i12 = LBound(arrWords(), 1) + 1 To UBound(arrWords(), 1) '+1 - pradedam nuo 1-o �od�io
            arrWords(i12, WordListID) = LCase(arrWords(i12, WordListID))
        Next i12
    End If
    
    ReDim arrTable(1 To TableSize, 1 To TableSize) As String   'Lentel� pagal �vest� lentel�s dyd�
    
    Dim i3 As Integer
    Dim arrWordsList() As String
    For i3 = LBound(arrWords, 1) + 1 To UBound(arrWords, 1) '�ra�o �od�ius � �onin� lente
        If FontCase = 1 Then
            shTable.Cells(i3, HelpWordsColumn) = arrWords(i3, WordListID) 'did�iosios
            Else: shTable.Cells(i3, HelpWordsColumn) = LCase(arrWords(i3, WordListID)) 'ma�osios
        End If
    Next i3
    
    Dim word As String                                      'einamasis �odis
    Dim i4 As Integer
    For i4 = LBound(arrWords, 1) + 1 To UBound(arrWords, 1)
        rndX = RndNum(TableSize)                            'randominam nauj� X
        rndY = RndNum(TableSize)                            'randominam nauj� Y
        
        Dim i5 As Integer
        Do Until IsEmpty(arrTable(rndX, rndY)) Or (i5 = 50) 'Jei elementas u�imtas - generuojam nauj� viet�
            i5 = i5 + 1
            rndX = RndNum(TableSize)                        'randominam nauj� X
            rndY = RndNum(TableSize)                        'randominam nauj� Y
        Loop
        
        WordLength = Len(arrWords(i4, WordListID))          '�od�io ilgis
        
        Fits = IsItFitIn(WordLength)                        'Tikrinam ar telpa � lentel�
        Fill Fits, WordLength, i4                     'Jei telpa - bandom �ra�yt arba kartojam
    Next i4
    
    If Not DEVELOPMODE Then Call ChrsToArrTable(arrChars()) '�ra�om � likusias vietas atsitiktines raides
    
    shTable.Protect                                         'U�rakinam shTable lap�
    
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
                IsItFitIn = 1                           'Telpa � vir��
                Else: IsItFitIn = -1                    'Netelpa � vir��
            End If
        Case 2                                          'RIGHT
            If (rndY + wordLen) <= TableSize + 1 Then
                IsItFitIn = 2                           'Telpa � de�in�
                Else: IsItFitIn = -2                    'Netelpa � de�in�
            End If
        Case 3                                          'DOWN
            If (rndX + wordLen) <= TableSize + 1 Then
                IsItFitIn = 3                           'Telpa � apa�i�
                Else: IsItFitIn = -3                    'Netelpa � apa�i�
            End If
        Case 4                                          'LEFT
            If (rndY - wordLen) >= 1 Then
                IsItFitIn = 4                           'Telpa � kair�
                Else: IsItFitIn = -4                    'Netelpa � kair�
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
                    Call DoItAgainBoy(i4)                       'Darom rekursij�
                    uzimta1 = True                              '�odis netelpa � vir��
                    Exit For
                End If
            Next i6
            If uzimta1 = False Then                             '�ra�om �od� � vir��
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
                    Call DoItAgainBoy(i4)                       'Darom rekursij�
                    uzimta2 = True                              '�odis netelpa � de�in�
                    Exit For
                End If
            Next i7
            If uzimta2 = False Then                             '�ra�om �od� � de�in�
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
                    Call DoItAgainBoy(i4)                       'Darom rekursij�
                    uzimta3 = True                              '�odis netelpa � apa�i�
                    Exit For
                End If
            Next i8
            If uzimta3 = False Then                             '�ra�om �od� � apa�i�
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
                    Call DoItAgainBoy(i4)                       'Darom rekursij�
                    uzimta4 = True                              '�odis nelepa � kair�
                    Exit For
                End If
            Next i9
            If uzimta4 = False Then                             '�ra�om �od� � kair�
                For i9 = 0 To WordLength - 1
                    Cells(rndX + 1, rndY + 1 - i9) = Mid(arrWords(i4, WordListID), i9 + 1, 1)
                    Call WriteCoordsToArray(i4)
                Next i9
            End If
        Case -1, -2, -3, -4                                     'Netelpa - i� naujo ie�kom vietos
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
        strArr(i, 1) = Mid(strChar, i, 1)       '1 masyvo dimensija - did�iosios raid�s
    Next i
    For i = 33 To 64                            'iteruojam strChar
        strArr(i - 32, 2) = Mid(strChar, i, 1)  '2 masyvo dimensija - ma�osios raid�s
    Next i
    PutChrsToArr = strArr()                     'gr��inam masyv�
End Function

Function ChrsToArrTable(ByRef arrChars() As String)
    '�dedam random raides � likusias masyvo ir lentel�s vietas
    Dim lrow As Integer
    Dim lcol As Integer
    Dim rndChar As Integer
    For lcol = LBound(arrTable, 2) To UBound(arrTable, 2)           'dvimatis masyvas
        For lrow = LBound(arrTable, 1) To UBound(arrTable, 1)       'vienmatis masyvas
            If (Cells(lcol + 1, lrow + 1) = "") Then                'jei cel� tu��ia
                rndChar = RndNum(32)                                'generuojam raid�
                arrTable(lcol, lrow) = arrChars(rndChar, FontCase)  '�ra�om raid� DID�/ma�
                Cells(lcol + 1, lrow + 1) = arrTable(lcol, lrow)    '�ra�om masyvo element� � lentel�
                Else: arrTable(lcol, lrow) = Cells(lcol + 1, lrow + 1)   '�ra�om cel� � masyv�
            End If
        Next lrow
    Next lcol
End Function

Function ClearTable()
    shWords.Range("rngWordsCheckBox1").Value = False    'Nuim� varnel� nuo CheckBox
    shTable.Unprotect                                   'Nuimam apsaug�
    shTable.Range("rngTableGrid").ClearContents
    shTable.Range("rngTableGrid").Interior.Color = xlNone
    shTable.Range("rngTableWords").ClearContents
    shTable.Range("rngTableWords").Interior.Color = xlNone
    shTable.Range("rngTableWords").Font.Strikethrough = False
    shTable.Range(Cells(2, 2), Cells(32, 32)).EntireRow.Hidden = False
    shTable.Range(Cells(2, 2), Cells(32, 32)).EntireColumn.Hidden = False
End Function

Function RndNum(ByRef n As Integer) As Integer
    Randomize                   'randominam skai�i�
    RndNum = Int(n * Rnd) + 1   'i�vedam randomint� skai�i�
End Function

Function HintMe()
    If DEVELOPMODE Then Exit Function                   'DEBUG
    If ActiveCell.Interior.Color <> 16777215 Then Exit Function 'Jei langelis spalvotas - i�eiti i� f-jos
    If ActiveCell.Column <> HelpWordsColumn Then Exit Function  'Jei nepa�ym�tas �od�i� stulpelis
    
    Dim x As Integer
    Dim y As Integer
    x = arrWords(ActiveCell.Row, arrWordsX)     'Nuskaitom eilut� kur prasida �odis
    y = arrWords(ActiveCell.Row, arrWordsY)     'Nuskaitom stulpel� kur prasideda �odis
    
    If Cells(x, y).Interior.Color <> 16777215 Then Exit Function
    
    shTable.Unprotect                          'Nuimam apsaug�
    Cells(x, y).Interior.Color = vbYellow      'Nuspalvinam
    shTable.UsedHintCount = shTable.UsedHintCount + 1
    shTable.Protect                             'U�dedam atgal apsaug�
End Function

Function WriteCoordsToArray(ByRef i4 As Integer)
    arrWords(i4, arrWordsX) = rndX + 1          '�ra�om kurioje eilut�je prasideda �odis
    arrWords(i4, arrWordsY) = rndY + 1          '�ra�om kuriame stulpelyje prasideda �odis
    arrWords(i4, WordDirection) = Fits          '� kuri� krypt� para�ytas �odis
End Function

Function DoItAgainBoy(ByRef i4 As Integer)
    rndX = RndNum(TableSize)        'Generuojam nauj� X
    rndY = RndNum(TableSize)        'Generuojam nauj� Y
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
