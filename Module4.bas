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
        .Value = Array("Vernr", "Bokföringsdatum", "Konto", "Benämning", "Ks", "Projnr", _
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
    Set dataRange = targetSheet.Range("Q" & startRow & ":Q" & lastRow)
    ReDim Preserve raderDennaMånad(0 To lastRow, 1)
    
    ' Läs in raderna
    raderDennaMånad = dataRange.Value
    
    'Skapa en variabel för att hålla den nuvarande verifikatsymbolen och en variabel för den föregående symbolen.
    Dim verifikatSymbol As String
    
    
    ' Definiera en boolean för att hålla reda på färgtoggel
    Dim color1bool As Boolean
    
    ' Förställ färg och symbol
    verifikatSymbol = raderDennaMånad(LBound(raderDennaMånad, 1), 1)
    color1bool = True
    j = 0
    Dim k As Long ' löpvariabel
    Dim rowCounter As Long
    
    rowCounter = 0
    For i = LBound(raderDennaMånad, 1) To UBound(raderDennaMånad, 1)
        If verifikatSymbol = raderDennaMånad(i, 1) Then
            ' Lägg till det globala Excelradnumret till verifikatRader
            verifikatRader(rowCounter) = startRow + i
            rowCounter = rowCounter + 1
        Else
            ' Nytt verifikat har hittats
            ' Färga raderna listade i verifikatRader från kolumn Q till AA med enligt colorBoolean
            For k = 0 To rowCounter - 1
                For Each cell In targetSheet.Range("Q" & verifikatRader(k) & ":AA" & verifikatRader(k))
                    cell.Interior.Color = IIf(color1bool, color1, color2)
                Next
            Next k
            
            ' Toggla colorBoolean
            color1bool = Not color1bool
            
            ' Töm verifikatRader arrayen och gör den redo för att ta emot nya rader
            Erase verifikatRader
            ReDim verifikatRader(100) ' Återställ storleken på vektorn
            rowCounter = 0
            
            ' Uppdatera verifikatSymbol till den nyhittade symbolen
            verifikatSymbol = raderDennaMånad(i, 1)
        End If
    Next i

         


End Sub
