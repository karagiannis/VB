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
    Dim k As Long ' löpvariabel
    Dim rowCounter As Long
    Dim rowIndex As Long
    Dim color1bool As Boolean ' Definiera en boolean för att hålla reda på färgtoggel
    Dim verifikatSymbol As String 'Skapa en variabel för att hålla den nuvarande verifikatsymbolen och en variabel för den föregående symbolen.
    
    
    
    Debug.Print "Inne i TestVerifikationslista()"
    ' Ange den aktuella arbetsboken
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Konvertera datumen i cellerna A2 och B2 till riktiga datumobjekt
    start_date = CDate(ws.Cells(2, 1).Value)
    Debug.Print start_date
    end_date = CDate(ws.Cells(2, 2).Value)
    Debug.Print end_date
    
    ' Extrahera månad från startdatumet
    currentMonth = month(start_date)
    
    ' Hämta förkortningen av månadens namn
    monthAbbreviation = Left(MonthName(currentMonth), 3)
    Debug.Print monthAbbreviation
    
    ' Hitta månadsfliken baserat på förkortningen av månadens namn
    On Error Resume Next ' Ignorera fel om månadsfliken inte finns
    Set targetSheet = ThisWorkbook.Sheets(monthAbbreviation)
    On Error GoTo 0 ' Återställ felhanteringen
    
    ' Kontrollera om månadsfliken hittades
    If Not targetSheet Is Nothing Then
        ' Gör något med månadsfliken
        Debug.Print "Månadsfliken " & monthAbbreviation & " hittades."
    Else
        ' Månadsfliken hittades inte
        Debug.Print "Månadsfliken " & monthAbbreviation & " hittades inte."
        ' exit
    End If
    
    
     'Aktivera fliken
    targetSheet.Activate
    
    ' Hitta den första och sista raden i verifikationslistan
    startRow = targetSheet.Cells(4, 7).Value
    lastRow = targetSheet.Cells(4, 8).Value
    Debug.Print "lastRow är: " & lastRow
    
    ' Ange området med data
    Set dataRange = targetSheet.Range("B" & startRow & ":B" & lastRow)
    ReDim Preserve data(startRow To lastRow, 1)
    
    ' Läs in datan till en edimensionell array
    data = dataRange.Value
    
    ' Skriv ut innehållet i variabeln data för att verifiera om den är tom eller inte
    Debug.Print "Innehållet i variabeln data:"
    For i = LBound(data, 1) To UBound(data, 1)
        Debug.Print "Rad " & i & ":" & data(i, 1)
    Next i
    
     'Kopiera rader från Kolumn A till K (1 till 11) för intervallet start_date till end_date till kolumn 16
    ' For i = LBound(data, 1) To UBound(data, 1)
    '   If data(i, 1) >= start_date And data(i, 1) <= end_date Then
    '      targetSheet.Rows(i).Range("A:F").Copy _
    '                Destination:=targetSheet.Rows(i).Offset(0, 15) ' 15 är kolumn P
    '
      '  End If
    '  Next i
    ' Dim col As Variant
    
    Dim headerRange As Range
    
    ' Ange området för rubrikerna
    Set headerRange = targetSheet.Range("Q" & startRow).Resize(1, 14)
    
    ' Skriv över bara rubrikerna och behåll befintliga värden i kolumnerna till höger
    With headerRange
        .Value = Array("Vernr", "Bokföringsdatum", "Registreringsdatum", "Konto", "Benämning", "Ks", "Projnr", _
                       "Verifikationstext", "Transaktionsinfo", "Debet", "Kredit", _
                       "Rätt moms", "Konto", "Aktiverad", "Har Flik")
        .Font.Bold = True
    End With



    For i = LBound(data, 1) To UBound(data, 1)
        If data(i, 1) >= start_date And data(i, 1) <= end_date Then
            For j = 1 To 11
             targetSheet.Cells(startRow + i - 1, j + 16) = targetSheet.Cells(startRow + i - 1, j)
            Next j
        End If
    Next i
    
    
    
    ' Ta reda på vilken rad månaden slutar
    Dim endOfMonthRow As Long
    ' Hitta den sista raden med data i kolumn Q (förutsatt att startRow är den första raden i sekvensen)
    endOfMonthRow = targetSheet.Cells(targetSheet.Rows.Count, "Q").End(xlUp).Row

    
    Dim lightBlue As Long
    lightBlue = RGB(200, 230, 255)
    
    Dim lightGreen As Long
    lightGreen = RGB(200, 255, 200)


    
    ' Definera en array för att hålla rad-data över en sammanhängande verifikatpost
    Dim verifikatRader() ' Deklarera en dynamisk array

    ' Ange storleken på den initiala dynamiska arrayen
    ReDim verifikatRader(100)
    
    'Definera behållarvariabel att löpa igenom
    Dim raderDennaMånad() As Variant
    
    ' Ange området med data
    Set dataRange = targetSheet.Range("Q" & startRow + 1 & ":Q" & lastRow)
    ReDim Preserve raderDennaMånad(1 To lastRow, 1)
    
    ' Läs in raderna
    raderDennaMånad = dataRange.Value
    Debug.Print LBound(raderDennaMånad, 1)
    
    ' Förställ färg och symbol
    verifikatSymbol = raderDennaMånad(1, 1)
    color1bool = True
    j = 0
    rowCounter = 0
    rowIndex = startRow + 1
    Debug.Print "rowIndex" & rowIndex
    Debug.Print "verifikatSymbol: " & verifikatSymbol
    
    ' Skapa en dictionary för att lagra radnummer för varje unikt verifikat
Dim verifikatRadnummer
Set verifikatRadnummer = CreateObject("Scripting.Dictionary")

' Förställ färg och symbol
verifikatSymbol = raderDennaMånad(1, 1)
color1bool = True
rowIndex = startRow + 1




i = 1 ' Börja från första raden
Do
    
    If verifikatSymbol = raderDennaMånad(i, 1) Then
        ' Lägg till det globala Excelradnumret till verifikatRader
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
        ' Färga raderna listade i verifikatRader från kolumn Q till AA med enligt colorBoolean
        For Each radnummer In verifikatRadnummer(verifikatSymbol)
            Debug.Print "radnummer :" & radnummer
            For Each Cell In targetSheet.Range("Q" & radnummer & ":AA" & radnummer)
                Cell.Interior.Color = IIf(color1bool, lightGreen, lightBlue)
            Next
        Next
        
        ' Toggla colorBoolean
        color1bool = Not color1bool
        
        ' Rensa dictionaryn för det nuvarande verifikatet
        verifikatRadnummer.Remove verifikatSymbol
        
        verifikatSymbol = raderDennaMånad(i, 1)
        i = i - 1
    End If
    
    ' Öka indexet för att gå till nästa rad
    i = i + 1

Loop Until i > UBound(raderDennaMånad, 1)
For Each radnummer In verifikatRadnummer(verifikatSymbol)
            Debug.Print "radnummer :" & radnummer
            For Each Cell In targetSheet.Range("Q" & radnummer & ":AA" & radnummer)
                Cell.Interior.Color = IIf(color1bool, lightGreen, lightBlue)
            Next
        Next
         

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Rensa dictionaryn för det nuvarande verifikatet
verifikatSymbol = raderDennaMånad(1, 1)
rowIndex = startRow + 1
i = 1 ' Börja från första raden
Do
    
    If verifikatSymbol = raderDennaMånad(i, 1) Then
        ' Lägg till det globala Excelradnumret till verifikatRader
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
        ' Färga raderna listade i verifikatRader från kolumn Q till AA med enligt colorBoolean
        For Each radnummer In verifikatRadnummer(verifikatSymbol)
            targetSheet.Range("AC" & radnummer).Value = targetSheet.Range("T" & radnummer).Value
        Next
        
       
        
        ' Rensa dictionaryn för det nuvarande verifikatet
        verifikatRadnummer.Remove verifikatSymbol
        
        verifikatSymbol = raderDennaMånad(i, 1)
        i = i - 1
    End If
    
    ' Öka indexet för att gå till nästa rad
    i = i + 1

Loop Until i > UBound(raderDennaMånad, 1)
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


' Rensa dictionaryn för det nuvarande verifikatet
verifikatSymbol = raderDennaMånad(1, 1)
rowIndex = startRow + 1
i = 1 ' Börja från första raden
Do
    
    If verifikatSymbol = raderDennaMånad(i, 1) Then
        ' Lägg till det globala Excelradnumret till verifikatRader
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
        
                ' Kontrollera om intäktskonto (3###) finns i kolumn T
            If targetSheet.Range("T" & radnummer).Value Like "3###" Then
                foundSell = True
                a = foundSell
            End If
            
            ' Kontrollera om moms 2611 finns i kolumn AA
            If targetSheet.Range("T" & radnummer).Value Like "2611" Then
                ' Om vi har hittat ett intäktskonto och moms 2611, markera detta som korrekt momsredovisning
                    ' Hittade moms 2611
                    foundSellVAT = True
                    b = foundSellVAT
            End If
            targetSheet.Range("AB" & radnummer).Value = "OK"
            
            ' Om moms 264# finns, markera detta som felaktig momsredovisning
            If targetSheet.Range("T" & radnummer).Value Like "264#" Then
                ' Hittade felaktig moms 264# som ska vara för inköp
                foundBuyVAT = True
                c = foundBuyVAT
            End If
            
            If ((a And b And Not (c)) Or (Not (a) And Not (b) And c) Or (Not (a) And b And c)) Then
                targetSheet.Range("AB" & radnummer).Value = "OK"
            Else
                targetSheet.Range("AB" & radnummer).Value = "NOK"
            End If
            
                
            
            
        Next
        
        ' Återställ variabler för nästa verifikat
        foundSell = False
        foundSellVAT = False
        foundBuyVAT = False
        a = False
        b = False
        c = False
        
        ' Rensa dictionaryn för det nuvarande verifikatet
        verifikatRadnummer.Remove verifikatSymbol
        
        ' Uppdatera till nästa verifikat
        verifikatSymbol = raderDennaMånad(i, 1)
        i = i - 1
    End If
    
    ' Öka indexet för att gå till nästa rad
    i = i + 1

Loop Until i > UBound(raderDennaMånad, 1)

' För den sista verifikatposten
For Each radnummer In verifikatRadnummer(verifikatSymbol)

               ' Kontrollera om intäktskonto (3###) finns i kolumn T
        If targetSheet.Range("T" & radnummer).Value Like "3###" Then
            foundSell = True
            a = foundSell
        End If
        
        ' Kontrollera om moms 2611 finns i kolumn AA
        If targetSheet.Range("T" & radnummer).Value Like "2611" Then
            ' Om vi har hittat ett intäktskonto och moms 2611, markera detta som korrekt momsredovisning
                ' Hittade moms 2611
                foundSellVAT = True
                b = foundSellVAT
        End If
        targetSheet.Range("AB" & radnummer).Value = "OK"
        
        ' Om moms 264# finns, markera detta som felaktig momsredovisning
        If targetSheet.Range("T" & radnummer).Value Like "264#" Then
            ' Hittade felaktig moms 264# som ska vara för inköp
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

' Deklarera variabler för tillstånd
Const start_state = 0
Const sell_state = 1
Const buy_moms_state = 2
Const sell_moms_state = 3
Const err_state = 4
Const OK_state = 5
Const momsrapport_found_state = 6

' Initialisera tillståndet
Dim state
state = start_state

' Loopa genom rader
For Each rad In verifikatRadnummer(verifikatSymbol)
    Select Case state
        Case start_state
            If rad Innehåller "3###" Then
                state = sell_state
            ElseIf rad Innehåller "264#" Then
                state = buy_moms_state
            ElseIf rad Innehåller "2611" Then
                state = sell_moms_state
            Else
                state = err_state
            End If
        Case sell_state
            If rad Innehåller "264#" Then
                state = err_state
            ElseIf rad Innehåller "2611" Then
                state = OK_state
            End If
        Case buy_moms_state
            If rad Innehåller "2611" Then
                state = momsrapport_found_state
            ElseIf rad Innehåller "264#" Then
                state = buy_moms_state
            ElseIf rad Innehåller "3###" Then
                state = err_state
            End If
        Case sell_moms_state
            If rad Innehåller "264#" Then
                state = momsrapport_found_state
            ElseIf rad Innehåller "3###" Then
                state = OK_state
            End If
    End Select
Next

