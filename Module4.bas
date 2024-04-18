Attribute VB_Name = "Module4"
Sub TestVerifikationslista()
Dim ws As Worksheet
    Dim start_date As Date
    Dim end_date As Date
    Dim currentMonth As Integer
    Dim monthAbbreviation As String
    Dim targetSheet As Worksheet
    Dim startRow As Long
    Dim lastRow As Long
    Dim datumIntervalStartingRow As Long
    Dim datumIntervalEndingRow As Long
    Dim i As Long, j As Long
    Dim dataRange As Range
    Dim data() As Variant
    Dim k As Long ' l�pvariabel
    Dim rowCounter As Long
    Dim rowIndex As Long
    Dim color1bool As Boolean ' Definiera en boolean f�r att h�lla reda p� f�rgtoggel
    Dim verifikatSymbol As String 'Skapa en variabel f�r att h�lla den nuvarande verifikatsymbolen och en variabel f�r den f�reg�ende symbolen.
    
    
    
    Debug.Print "Inne i TestVerifikationslista()"
    ' Ange den aktuella arbetsboken
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Konvertera datumen i cellerna A2 och B2 till riktiga datumobjekt
    start_date = CDate(ws.Cells(2, 1).Value)
    Debug.Print start_date
    end_date = CDate(ws.Cells(2, 2).Value)
    Debug.Print end_date
    
    ' Extrahera m�nad fr�n startdatumet
    currentMonth = month(start_date)
    
    ' H�mta f�rkortningen av m�nadens namn
    monthAbbreviation = Left(MonthName(currentMonth), 3)
    Debug.Print monthAbbreviation
    
    ' Hitta m�nadsfliken baserat p� f�rkortningen av m�nadens namn
    On Error Resume Next ' Ignorera fel om m�nadsfliken inte finns
    Set targetSheet = ThisWorkbook.Sheets(monthAbbreviation)
    On Error GoTo 0 ' �terst�ll felhanteringen
    
    ' Kontrollera om m�nadsfliken hittades
    If Not targetSheet Is Nothing Then
        ' G�r n�got med m�nadsfliken
        Debug.Print "M�nadsfliken " & monthAbbreviation & " hittades."
    Else
        ' M�nadsfliken hittades inte
        Debug.Print "M�nadsfliken " & monthAbbreviation & " hittades inte."
        ' exit
    End If
    
    
     'Aktivera fliken
    targetSheet.Activate
    
    ' Hitta den f�rsta och sista raden i verifikationslistan
    startRow = targetSheet.Cells(4, 7).Value
    lastRow = targetSheet.Cells(4, 8).Value
    Debug.Print "lastRow �r: " & lastRow
    
    ' Ange omr�det med data
    Set dataRange = targetSheet.Range("B" & startRow & ":B" & lastRow)
    ReDim Preserve data(startRow To lastRow, 1)
    
    ' L�s in datan till en edimensionell array
    data = dataRange.Value
    
    ' Skriv ut inneh�llet i variabeln data f�r att verifiera om den �r tom eller inte
    Debug.Print "Inneh�llet i variabeln data:"
    For i = LBound(data, 1) To UBound(data, 1)
        Debug.Print "Rad " & i & ":" & data(i, 1)
    Next i
    
     'Kopiera rader fr�n Kolumn A till K (1 till 11) f�r intervallet start_date till end_date till kolumn 16
    ' For i = LBound(data, 1) To UBound(data, 1)
    '   If data(i, 1) >= start_date And data(i, 1) <= end_date Then
    '      targetSheet.Rows(i).Range("A:F").Copy _
    '                Destination:=targetSheet.Rows(i).Offset(0, 15) ' 15 �r kolumn P
    '
      '  End If
    '  Next i
    ' Dim col As Variant
    
    Dim headerRange As Range
    
    ' Ange omr�det f�r rubrikerna
    Set headerRange = targetSheet.Range("Q" & startRow).Resize(1, 14)
    
    ' Skriv �ver bara rubrikerna och beh�ll befintliga v�rden i kolumnerna till h�ger
    With headerRange
        .Value = Array("Vernr", "Bokf�ringsdatum", "Registreringsdatum", "Konto", "Ben�mning", "Ks", "Projnr", _
                       "Verifikationstext", "Transaktionsinfo", "Debet", "Kredit", _
                       "R�tt moms", "Konto", "Aktiverad", "Har Flik")
        .Font.Bold = True
    End With



    For i = LBound(data, 1) To UBound(data, 1)
        If data(i, 1) >= start_date And data(i, 1) <= end_date Then
            For j = 1 To 11
             targetSheet.Cells(startRow + i - 1, j + 16) = targetSheet.Cells(startRow + i - 1, j)
            Next j
        End If
    Next i
    
    
    
    ' Ta reda p� vilken rad m�naden slutar
    Dim endOfMonthRow As Long
    ' Hitta den sista raden med data i kolumn Q (f�rutsatt att startRow �r den f�rsta raden i sekvensen)
    endOfMonthRow = targetSheet.Cells(targetSheet.Rows.Count, "Q").End(xlUp).Row

    
    Dim lightBlue As Long
    lightBlue = RGB(200, 230, 255)
    
    Dim lightGreen As Long
    lightGreen = RGB(200, 255, 200)


    
    ' Definera en array f�r att h�lla rad-data �ver en sammanh�ngande verifikatpost
    Dim verifikatRader() ' Deklarera en dynamisk array

    ' Ange storleken p� den initiala dynamiska arrayen
    ReDim verifikatRader(100)
    
    'Definera beh�llarvariabel att l�pa igenom
    Dim raderDennaM�nad() As Variant
    
    ' Ange omr�det med data
    Set dataRange = targetSheet.Range("Q" & startRow + 1 & ":Q" & lastRow)
    ReDim Preserve raderDennaM�nad(1 To lastRow, 1)
    
    ' L�s in raderna
    raderDennaM�nad = dataRange.Value
    Debug.Print LBound(raderDennaM�nad, 1)
    
    ' F�rst�ll f�rg och symbol
    verifikatSymbol = raderDennaM�nad(1, 1)
    color1bool = True
    j = 0
    rowCounter = 0
    rowIndex = startRow + 1
    Debug.Print "rowIndex" & rowIndex
    Debug.Print "verifikatSymbol: " & verifikatSymbol
    
    ' Skapa en dictionary f�r att lagra radnummer f�r varje unikt verifikat
Dim verifikatRadnummer
Set verifikatRadnummer = CreateObject("Scripting.Dictionary")

' F�rst�ll f�rg och symbol
verifikatSymbol = raderDennaM�nad(1, 1)
color1bool = True
rowIndex = startRow + 1




i = 1 ' B�rja fr�n f�rsta raden
Do
    
    If verifikatSymbol = raderDennaM�nad(i, 1) Then
        ' L�gg till det globala Excelradnumret till verifikatRader
        Debug.Print "Before adding rowIndex:", rowIndex
        If Not verifikatRadnummer.Exists(verifikatSymbol) Then
            verifikatRadnummer.Add verifikatSymbol, New Collection
        End If
        Debug.Print "Number of elements in Collection:", verifikatRadnummer(verifikatSymbol).Count
        verifikatRadnummer(verifikatSymbol).Add rowIndex
        Debug.Print "After adding rowIndex:", rowIndex
        Debug.Print "Number of elements in Collection:", verifikatRadnummer(verifikatSymbol).Count
        rowIndex = rowIndex + 1
    Else
        ' Nytt verifikat har hittats
        ' F�rga raderna listade i verifikatRader fr�n kolumn Q till AA med enligt colorBoolean
        For Each radnummer In verifikatRadnummer(verifikatSymbol)
            Debug.Print "radnummer :" & radnummer
            For Each Cell In targetSheet.Range("Q" & radnummer & ":AA" & radnummer)
                Cell.Interior.Color = IIf(color1bool, lightGreen, lightBlue)
            Next
        Next
        
        ' Toggla colorBoolean
        color1bool = Not color1bool
        
        ' Rensa dictionaryn f�r det nuvarande verifikatet
        verifikatRadnummer.Remove verifikatSymbol
        
        verifikatSymbol = raderDennaM�nad(i, 1)
        i = i - 1
    End If
    
    ' �ka indexet f�r att g� till n�sta rad
    i = i + 1

Loop Until i > UBound(raderDennaM�nad, 1)
For Each radnummer In verifikatRadnummer(verifikatSymbol)
            Debug.Print "radnummer :" & radnummer
            For Each Cell In targetSheet.Range("Q" & radnummer & ":AA" & radnummer)
                Cell.Interior.Color = IIf(color1bool, lightGreen, lightBlue)
            Next
        Next
         

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Rensa dictionaryn f�r det nuvarande verifikatet
verifikatSymbol = raderDennaM�nad(1, 1)
rowIndex = startRow + 1
i = 1 ' B�rja fr�n f�rsta raden
Do
    
    If verifikatSymbol = raderDennaM�nad(i, 1) Then
        ' L�gg till det globala Excelradnumret till verifikatRader
        Debug.Print "Before adding rowIndex:", rowIndex
        If Not verifikatRadnummer.Exists(verifikatSymbol) Then
            verifikatRadnummer.Add verifikatSymbol, New Collection
        End If
        Debug.Print "Number of elements in Collection:", verifikatRadnummer(verifikatSymbol).Count
        verifikatRadnummer(verifikatSymbol).Add rowIndex
        Debug.Print "After adding rowIndex:", rowIndex
        Debug.Print "Number of elements in Collection:", verifikatRadnummer(verifikatSymbol).Count
        rowIndex = rowIndex + 1
    Else
        ' Nytt verifikat har hittats
        ' F�rga raderna listade i verifikatRader fr�n kolumn Q till AA med enligt colorBoolean
        For Each radnummer In verifikatRadnummer(verifikatSymbol)
            targetSheet.Range("AC" & radnummer).Value = targetSheet.Range("T" & radnummer).Value
        Next
        
       
        
        ' Rensa dictionaryn f�r det nuvarande verifikatet
        verifikatRadnummer.Remove verifikatSymbol
        
        verifikatSymbol = raderDennaM�nad(i, 1)
        i = i - 1
    End If
    
    ' �ka indexet f�r att g� till n�sta rad
    i = i + 1

Loop Until i > UBound(raderDennaM�nad, 1)
 For Each radnummer In verifikatRadnummer(verifikatSymbol)
        targetSheet.Range("AC" & radnummer).Value = targetSheet.Range("T" & radnummer).Value
    Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim sellAccount As Long
Dim vatAccount As Long
Dim foundSell As Boolean
Dim foundVAT As Boolean
Dim foundVATcodeError As Boolean
foundSell = False
foundSellVAT = False
foundBuyVAT = False
Dim a As Boolean, b As Boolean, c As Boolean
a = False
b = False
c = False


' Rensa dictionaryn f�r det nuvarande verifikatet
verifikatSymbol = raderDennaM�nad(1, 1)
rowIndex = startRow + 1
i = 1 ' B�rja fr�n f�rsta raden
Do
    
    If verifikatSymbol = raderDennaM�nad(i, 1) Then
        ' L�gg till det globala Excelradnumret till verifikatRader
        Debug.Print "Before adding rowIndex:", rowIndex
        If Not verifikatRadnummer.Exists(verifikatSymbol) Then
            verifikatRadnummer.Add verifikatSymbol, New Collection
        End If
        Debug.Print "Number of elements in Collection:", verifikatRadnummer(verifikatSymbol).Count
        verifikatRadnummer(verifikatSymbol).Add rowIndex
        Debug.Print "After adding rowIndex:", rowIndex
        Debug.Print "Number of elements in Collection:", verifikatRadnummer(verifikatSymbol).Count
        rowIndex = rowIndex + 1
    Else
        ' Nytt verifikat har hittats
        For Each radnummer In verifikatRadnummer(verifikatSymbol)
        
                ' Kontrollera om int�ktskonto (3###) finns i kolumn T
            If targetSheet.Range("T" & radnummer).Value Like "3###" Then
                foundSell = True
                a = foundSell
            End If
            
            ' Kontrollera om moms 2611 finns i kolumn AA
            If targetSheet.Range("T" & radnummer).Value Like "2611" Then
                ' Om vi har hittat ett int�ktskonto och moms 2611, markera detta som korrekt momsredovisning
                    ' Hittade moms 2611
                    foundSellVAT = True
                    b = foundSellVAT
            End If
            targetSheet.Range("AB" & radnummer).Value = "OK"
            
            ' Om moms 264# finns, markera detta som felaktig momsredovisning
            If targetSheet.Range("T" & radnummer).Value Like "264#" Then
                ' Hittade felaktig moms 264# som ska vara f�r ink�p
                foundBuyVAT = True
                c = foundBuyVAT
            End If
            
            If ((a And b And Not (c)) Or (Not (a) And Not (b) And c) Or (Not (a) And b And c)) Then
                targetSheet.Range("AB" & radnummer).Value = "OK"
            Else
                targetSheet.Range("AB" & radnummer).Value = "NOK"
            End If
            
                
            
            
        Next
        
        ' �terst�ll variabler f�r n�sta verifikat
        foundSell = False
        foundSellVAT = False
        foundBuyVAT = False
        a = False
        b = False
        c = False
        
        ' Rensa dictionaryn f�r det nuvarande verifikatet
        verifikatRadnummer.Remove verifikatSymbol
        
        ' Uppdatera till n�sta verifikat
        verifikatSymbol = raderDennaM�nad(i, 1)
        i = i - 1
    End If
    
    ' �ka indexet f�r att g� till n�sta rad
    i = i + 1

Loop Until i > UBound(raderDennaM�nad, 1)

' F�r den sista verifikatposten
For Each radnummer In verifikatRadnummer(verifikatSymbol)

               ' Kontrollera om int�ktskonto (3###) finns i kolumn T
        If targetSheet.Range("T" & radnummer).Value Like "3###" Then
            foundSell = True
            a = foundSell
        End If
        
        ' Kontrollera om moms 2611 finns i kolumn AA
        If targetSheet.Range("T" & radnummer).Value Like "2611" Then
            ' Om vi har hittat ett int�ktskonto och moms 2611, markera detta som korrekt momsredovisning
                ' Hittade moms 2611
                foundSellVAT = True
                b = foundSellVAT
        End If
        targetSheet.Range("AB" & radnummer).Value = "OK"
        
        ' Om moms 264# finns, markera detta som felaktig momsredovisning
        If targetSheet.Range("T" & radnummer).Value Like "264#" Then
            ' Hittade felaktig moms 264# som ska vara f�r ink�p
            foundBuyVAT = True
            c = foundBuyVAT
        End If
        
        If ((a And b And Not (c)) Or (Not (a) And Not (b) And c) Or (Not (a) And b And c)) Then
            targetSheet.Range("AB" & radnummer).Value = "OK"
        Else
            targetSheet.Range("AB" & radnummer).Value = "NOK"
        End If
            

Next

End Sub

' Deklarera variabler f�r tillst�nd
Const start_state = 0
Const sell_state = 1
Const buy_moms_state = 2
Const sell_moms_state = 3
Const err_state = 4
Const OK_state = 5
Const momsrapport_found_state = 6

' Initialisera tillst�ndet
Dim state
state = start_state

' Loopa genom rader
For Each rad In verifikatRadnummer(verifikatSymbol)
    Select Case state
        Case start_state
            If rad Inneh�ller "3###" Then
                state = sell_state
            ElseIf rad Inneh�ller "264#" Then
                state = buy_moms_state
            ElseIf rad Inneh�ller "2611" Then
                state = sell_moms_state
            Else
                state = err_state
            End If
        Case sell_state
            If rad Inneh�ller "264#" Then
                state = err_state
            ElseIf rad Inneh�ller "2611" Then
                state = OK_state
            End If
        Case buy_moms_state
            If rad Inneh�ller "2611" Then
                state = momsrapport_found_state
            ElseIf rad Inneh�ller "264#" Then
                state = buy_moms_state
            ElseIf rad Inneh�ller "3###" Then
                state = err_state
            End If
        Case sell_moms_state
            If rad Inneh�ller "264#" Then
                state = momsrapport_found_state
            ElseIf rad Inneh�ller "3###" Then
                state = OK_state
            End If
    End Select
Next

