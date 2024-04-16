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
        .Value = Array("Vernr", "Bokf�ringsdatum", "Konto", "Ben�mning", "Ks", "Projnr", _
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
    Set dataRange = targetSheet.Range("Q" & startRow & ":Q" & lastRow)
    ReDim Preserve raderDennaM�nad(0 To lastRow, 1)
    
    ' L�s in raderna
    raderDennaM�nad = dataRange.Value
    
    'Skapa en variabel f�r att h�lla den nuvarande verifikatsymbolen och en variabel f�r den f�reg�ende symbolen.
    Dim verifikatSymbol As String
    
    
    ' Definiera en boolean f�r att h�lla reda p� f�rgtoggel
    Dim color1bool As Boolean
    
    ' F�rst�ll f�rg och symbol
    verifikatSymbol = raderDennaM�nad(LBound(raderDennaM�nad, 1), 1)
    color1bool = True
    j = 0
    Dim k As Long ' l�pvariabel
    Dim rowCounter As Long
    
    rowCounter = 0
    For i = LBound(raderDennaM�nad, 1) To UBound(raderDennaM�nad, 1)
        If verifikatSymbol = raderDennaM�nad(i, 1) Then
            ' L�gg till det globala Excelradnumret till verifikatRader
            verifikatRader(rowCounter) = startRow + i
            rowCounter = rowCounter + 1
        Else
            ' Nytt verifikat har hittats
            ' F�rga raderna listade i verifikatRader fr�n kolumn Q till AA med enligt colorBoolean
            For k = 0 To rowCounter - 1
                For Each cell In targetSheet.Range("Q" & verifikatRader(k) & ":AA" & verifikatRader(k))
                    cell.Interior.Color = IIf(color1bool, color1, color2)
                Next
            Next k
            
            ' Toggla colorBoolean
            color1bool = Not color1bool
            
            ' T�m verifikatRader arrayen och g�r den redo f�r att ta emot nya rader
            Erase verifikatRader
            ReDim verifikatRader(100) ' �terst�ll storleken p� vektorn
            rowCounter = 0
            
            ' Uppdatera verifikatSymbol till den nyhittade symbolen
            verifikatSymbol = raderDennaM�nad(i, 1)
        End If
    Next i

         


End Sub
