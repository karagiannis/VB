Attribute VB_Name = "Module14"
Sub LäsInBalansrapport()
    Dim filnamn As String
    Dim rot_mapp As String
    Dim mapp As String
    Dim bladNamn As String
    Dim kolumnIndex As Integer
    Dim ws As Worksheet
    Dim filnamnsArray() As String
    Dim balansArray() As String
    Dim xlsxArray() As String
    Dim i As Integer
    Dim j As Integer
    Dim start_date As Date
    Dim end_date As Date
    
    
    ' Ange den aktuella arbetsboken
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Ange mappen där balansrapporten finns
    rot_mapp = "G:\Bokföring\Planering inför Årsbokslut"
    
    ' Läs in namnet på den månadsmapp som balansrapporten ska hämtas ifrån
    mapp = ws.Cells(1, 1).Value
    
    ' Skriv ut sökvägen till mappen
    Debug.Print rot_mapp & "\" & mapp
    
    ' Återställ arrayen för att undvika eventuellt tidigare innehåll
    ReDim filnamnsArray(0)
    
    ' Läs in filnamnen som ligger i mappen och lägg filnamnen i en array
    filnamn = Dir(rot_mapp & "\" & mapp & "\")
    i = 0
    Do While filnamn <> "" ' Loopa så länge det finns filnamn i mappen
        Debug.Print filnamn
        ' Lägg till filnamnet i arrayen
        ReDim Preserve filnamnsArray(i)
        filnamnsArray(i) = filnamn
        i = i + 1
        ' Läs nästa filnamn
        filnamn = Dir
    Loop
    
    i = 0
    j = 0
    Debug.Print "After while loop completed"
    
    ' Sök efter filnamn som börjar med "Balans" och lägg dem i en annan array
    For j = LBound(filnamnsArray) To UBound(filnamnsArray)
        If Left(filnamnsArray(j), 6) = "Balans" Then
            ReDim Preserve balansArray(i)
            balansArray(i) = filnamnsArray(j)
            i = i + 1
        End If
    Next j
     Debug.Print "Balansarray filled"
    ' Skriv ut alla filnamn som börjar med "Balans" i arrayen
     For i = LBound(balansArray) To UBound(balansArray)
       Debug.Print balansArray(i)
     Next i
     
    i = 0
    j = 0
    Debug.Print "Before xlsarray filling"
    ' Sök efter filnamn som slutar med ".xlsx" och lägg dem i en annan array
    For j = LBound(balansArray) To UBound(balansArray)
        If Right(balansArray(j), 5) = ".xlsx" Then
            ReDim Preserve xlsxArray(i)
            xlsxArray(i) = balansArray(j)
            i = i + 1
        End If
    Next j
    Debug.Print "xlsarray filled"
    ' Skriv ut alla filnamn som slutar med ".xlsx" i arrayen
    For i = LBound(xlsxArray) To UBound(xlsxArray)
       Debug.Print xlsxArray(i)
     Next i
    
    ' Konvertera datumen i cellerna A2 och B2 till riktiga datumobjekt
    start_date = CDate(ws.Cells(2, 1).Value)
    Debug.Print start_date
    end_date = CDate(ws.Cells(2, 2).Value)
    Debug.Print end_date

    ' Skriv ut alla filnamn som uppfyller villkoren i arrayen
    Dim parts() As String
    Dim hittad As Boolean
    Dim date_parts() As String
    hittad = False
    Dim file_start_date As Date
    Dim file_end_date As Date
    Dim rätt_balansrapport As String
    
    Dim year_start As Integer
    Dim month_start As Integer
    Dim day_start As Integer
    Dim year_end As Integer
    Dim month_end As Integer
    Dim day_end As Integer
    
    For i = LBound(xlsxArray) To UBound(xlsxArray)
        Debug.Print xlsxArray(i)
        
        ' Splitta filnamnet vid underscore-tecknet
        
        parts = Split(xlsxArray(i), "_")
        Debug.Print parts(1)
        
        ' Kontrollera att det finns minst två delar efter uppdelning
        If UBound(parts) >= 1 Then
            ' Splitta den andra delen vid minustecknet
            date_parts = Split(parts(1), "-")
            Debug.Print date_parts(0)
            Debug.Print date_parts(1)
            
        End If
        ' Kontrollera att det finns exakt två delar efter uppdelning
        If UBound(date_parts) = 1 Then
            ' Extrahera start- och slutdatum från filnamnet
            ' Dela upp start- och slutdatumet i år, månad och dag
            year_start = Left(date_parts(0), 4)
            month_start = Mid(date_parts(0), 5, 2)
            day_start = Right(date_parts(0), 2)

            year_end = Left(date_parts(1), 4)
            month_end = Mid(date_parts(1), 5, 2)
            day_end = Right(date_parts(1), 2)

            ' Skapa start- och slutdatumet med DateSerial
            file_start_date = DateSerial(year_start, month_start, day_start)
            file_end_date = DateSerial(year_end, month_end, day_end)
            
            ' Kontrollera om filens datumintervall matchar det angivna intervallet
            If file_start_date = start_date And file_end_date = end_date Then
                ' Sätt flaggan och lagra indexet för det matchande filnamnet
                rätt_balansrapport = xlsxArray(i)
                hittad = True
                Exit For ' Avbryt loopen när rätt filnamn har hittats
            End If
        End If
    Next i
    
    If hittad Then
        ' Skriv ut det valda filnamnet
        Debug.Print "Rätt balansrapport: " & rätt_balansrapport
    Else
        ' Filen hittas inte
        Debug.Print "Rätt balansrapport har inte hittats"
    End If
    
    Dim monthAbbreviation
    monthAbbreviation = Left(MonthName(month_start), 3)

    Debug.Print monthAbbreviation
    
    Dim targetSheet As Worksheet
    Dim sheetName As String
    sheetName = monthAbbreviation ' Använd innehållet i monthAbbreviation som fliknamn

    ' Kontrollera om fliken med det aktuella namnet redan finns
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' Om fliken inte finns, skapa den
    If targetSheet Is Nothing Then
        Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        targetSheet.Name = sheetName
    End If
    
    
    
    ' Skapa sökvägen till filen
    Dim path As String
    path = rot_mapp & "\" & mapp & "\" & rätt_balansrapport
    Debug.Print "Sökvägen till balansrapporten är:"
    Debug.Print path
    
    
    ' Öppna den andra filen
    Dim balansWorkbook As Workbook
    Set balansWorkbook = Workbooks.Open(path)

   
    targetSheet.Activate
    TaBortKolumner targetSheet

    ' Kopiera innehållet från den enda fliken i balansWorkbook till targetSheet
    balansWorkbook.Sheets(1).UsedRange.Copy targetSheet.Range("A1")

    ' Stäng den andra filen utan att spara ändringar
    balansWorkbook.Close SaveChanges:=False

    JusteraKolumnBredden targetSheet
    
    InfogaHeaders targetSheet
    
    targetSheet.Cells(1, 7).Value = file_start_date
    targetSheet.Cells(1, 8).Value = file_end_date
    
    LäsInResultatrapport targetSheet, rot_mapp, mapp, file_start_date, file_end_date
    LäsInHuvudbok targetSheet, rot_mapp, mapp, file_start_date, file_end_date
    LäsInVerifikationslista targetSheet, rot_mapp, mapp, file_start_date, file_end_date
    Debug.Print "test"
End Sub
Sub TaBortKolumner(ws As Worksheet)
    ' Ta bort kolumnerna A till J
   ws.Columns("A:J").Clear
End Sub


Sub JusteraKolumnBredden(ByVal targetSheet As Worksheet)
    Dim columnRange As Range
    Dim i As Integer
    
    ' Ange vilka kolumner du vill justera bredden för
    ' Här antar jag att du vill justera alla kolumner från A till Z
    Set columnRange = targetSheet.Range("A:K")
    
    ' Ange den önskade bredden i tecken
    Dim desiredWidth As Integer
    desiredWidth = 15 ' Du kan ändra detta värde till det du anser passar bäst
    
    ' Justera bredden för varje kolumn i den angivna kolumnräckan
    For i = 1 To columnRange.Columns.Count
        columnRange.Columns(i).ColumnWidth = desiredWidth
    Next i
End Sub

Sub InfogaHeaders(ByVal targetSheet As Worksheet)
    Dim startRow As Long
    Dim lastRow As Long
    Dim headerRange As Range
    
    ' Ange startraden för att söka efter kriterierna
    startRow = 1
    
    ' Hitta den sista raden med data i första kolumnen
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    
    ' Sök igenom kolumn A för att hitta de specificerade kriterierna och infoga kolumnheaders
    For i = startRow To lastRow
        If targetSheet.Cells(i, 1).Value = "Materiella anläggningstillgångar" Or _
           targetSheet.Cells(i, 1).Value = "Kortfristiga fordringar" Or _
           targetSheet.Cells(i, 1).Value = "EGET KAPITAL, AVSÄTTNINGAR OCH SKULDER" Or _
           targetSheet.Cells(i, 1).Value = "Långfristiga skulder" Or _
           targetSheet.Cells(i, 1).Value = "Kortfristiga skulder" Then
           
           ' Infoga kolumnheaders med fetstil från kolumn C
           Set headerRange = targetSheet.Range("C" & i)
           headerRange.Resize(1, 17).Value = Array("Ing balans", "Ing saldo", "Period", _
                                                   "Utg balans", "Period beräknad", _
                                                   "Utg balans beräknad", "Överensstämmer", _
                                                   "Beräkningsunderlag", "1", "2", "3", "4", "5", "6", "7", "IB koll", "Saldo koll")
           headerRange.Resize(1, 17).Font.Bold = True
           
           
        End If
    Next i
End Sub

Sub LäsInResultatrapport(ByVal targetSheet As Worksheet, ByVal rot_mapp As String, ByVal mapp As String, ByVal start_date As Date, ByVal end_date As Date)
    Dim filnamn As String
    Dim bladNamn As String
    Dim kolumnIndex As Integer
    Dim filnamnsArray() As String
    Dim resultatArray() As String
    Dim xlsxArray() As String
    Dim i As Integer
    Dim j As Integer

    
    
    
    ' Skriv ut sökvägen till mappen
    Debug.Print rot_mapp & "\" & mapp
    
    ' Återställ arrayen för att undvika eventuellt tidigare innehåll
    ReDim filnamnsArray(0)
    
    ' Läs in filnamnen som ligger i mappen och lägg filnamnen i en array
    filnamn = Dir(rot_mapp & "\" & mapp & "\")
    i = 0
    Do While filnamn <> "" ' Loopa så länge det finns filnamn i mappen
        Debug.Print filnamn
        ' Lägg till filnamnet i arrayen
        ReDim Preserve filnamnsArray(i)
        filnamnsArray(i) = filnamn
        i = i + 1
        ' Läs nästa filnamn
        filnamn = Dir
    Loop
    
    i = 0
    j = 0
    Debug.Print "After while loop completed"
    
    ' Sök efter filnamn som börjar med "Balans" och lägg dem i en annan array
    For j = LBound(filnamnsArray) To UBound(filnamnsArray)
        If Left(filnamnsArray(j), 8) = "Resultat" Then
            ReDim Preserve resultatArray(i)
            resultatArray(i) = filnamnsArray(j)
            i = i + 1
        End If
    Next j
     Debug.Print "ResultatArray filled"
    ' Skriv ut alla filnamn som börjar med "Resultat" i arrayen
     If Not IsEmpty(resultatArray) Then
        For i = LBound(resultatArray) To UBound(resultatArray)
            Debug.Print resultatArray(i)
        Next i
    Else
        Debug.Print "Inga resultatfiler hittades."
    End If

     
    i = 0
    j = 0
    Debug.Print "Before xlsarray filling"
    ' Sök efter filnamn som slutar med ".xlsx" och lägg dem i en annan array
    For j = LBound(resultatArray) To UBound(resultatArray)
        If Right(resultatArray(j), 5) = ".xlsx" Then
            ReDim Preserve xlsxArray(i)
            xlsxArray(i) = resultatArray(j)
            i = i + 1
        End If
    Next j
    Debug.Print "xlsarray filled"
    ' Skriv ut alla filnamn som slutar med ".xlsx" i arrayen
    For i = LBound(xlsxArray) To UBound(xlsxArray)
       Debug.Print xlsxArray(i)
     Next i
    


    ' Skriv ut alla filnamn som uppfyller villkoren i arrayen
    Dim parts() As String
    Dim hittad As Boolean
    Dim date_parts() As String
    hittad = False
    Dim file_start_date As Date
    Dim file_end_date As Date
    Dim rätt_resultatrapport As String
    
    Dim year_start As Integer
    Dim month_start As Integer
    Dim day_start As Integer
    Dim year_end As Integer
    Dim month_end As Integer
    Dim day_end As Integer
    
    For i = LBound(xlsxArray) To UBound(xlsxArray)
        Debug.Print xlsxArray(i)
        
        ' Splitta filnamnet vid underscore-tecknet
        
        parts = Split(xlsxArray(i), "_")
        Debug.Print parts(1)
        
        ' Kontrollera att det finns minst två delar efter uppdelning
        If UBound(parts) >= 1 Then
            ' Splitta den andra delen vid minustecknet
            date_parts = Split(parts(1), "-")
            Debug.Print date_parts(0)
            Debug.Print date_parts(1)
            
        End If
        ' Kontrollera att det finns exakt två delar efter uppdelning
        If UBound(date_parts) = 1 Then
            ' Extrahera start- och slutdatum från filnamnet
            ' Dela upp start- och slutdatumet i år, månad och dag
            year_start = Left(date_parts(0), 4)
            month_start = Mid(date_parts(0), 5, 2)
            day_start = Right(date_parts(0), 2)

            year_end = Left(date_parts(1), 4)
            month_end = Mid(date_parts(1), 5, 2)
            day_end = Right(date_parts(1), 2)

            ' Skapa start- och slutdatumet med DateSerial
            file_start_date = DateSerial(year_start, month_start, day_start)
            file_end_date = DateSerial(year_end, month_end, day_end)
            
            ' Kontrollera om filens datumintervall matchar det angivna intervallet
            If file_start_date = start_date And file_end_date = end_date Then
                ' Sätt flaggan och lagra indexet för det matchande filnamnet
                rätt_resultatrapport = xlsxArray(i)
                hittad = True
                Exit For ' Avbryt loopen när rätt filnamn har hittats
            End If
        End If
    Next i
    
    If hittad Then
        ' Skriv ut det valda filnamnet
        Debug.Print "Rätt resultatrapport: " & rätt_resultatrapport
    Else
        ' Filen hittas inte
        Debug.Print "Rätt resultatrapport har inte hittats"
    End If
    
    Dim monthAbbreviation
    monthAbbreviation = Left(MonthName(month_start), 3)

    Debug.Print monthAbbreviation
  
    
    ' Skapa sökvägen till filen
    Dim path As String
    path = rot_mapp & "\" & mapp & "\" & rätt_resultatrapport
    Debug.Print "Sökvägen till resultatrapporten är:"
    Debug.Print path
    
    
    ' Öppna resultarrapporten
    Dim resultatWorkbook As Workbook
    Set resultatWorkbook = Workbooks.Open(path)
    Dim resultatSheet As Worksheet
    Set resultatSheet = resultatWorkbook.Sheets(1) ' Antag att resultatrapporten är på första arket

   
    targetSheet.Activate
    
    
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    
    ' Beräkna var du ska börja kopiera resultatrapporten
    Dim startRow As Long
    startRow = lastRow + 5 ' 5 rader under den sista raden i balansrapporten
    
    
    ' Kopiera resultatrapporten till det angivna området
    resultatSheet.UsedRange.Copy targetSheet.Cells(startRow, 1) ' Börja i den första kolumnen på startRow
    
    ' Stäng resultatrapporten utan att spara ändringar
    resultatWorkbook.Close SaveChanges:=False
    
    targetSheet.Cells(2, 7).Value = startRow
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    targetSheet.Cells(2, 8).Value = lastRow
    targetSheet.Cells(2, 6).Value = "Resultatrapport"
    
    

    Debug.Print "test"
End Sub


Sub LäsInHuvudbok(ByVal targetSheet As Worksheet, ByVal rot_mapp As String, ByVal mapp As String, ByVal start_date As Date, ByVal end_date As Date)
    Dim filnamn As String
    Dim bladNamn As String
    Dim kolumnIndex As Integer
    Dim filnamnsArray() As String
    Dim huvudbokArray() As String
    Dim xlsxArray() As String
    Dim i As Integer
    Dim j As Integer

    
    
    
    ' Skriv ut sökvägen till mappen
    Debug.Print rot_mapp & "\" & mapp
    
    ' Återställ arrayen för att undvika eventuellt tidigare innehåll
    ReDim filnamnsArray(0)
    
    ' Läs in filnamnen som ligger i mappen och lägg filnamnen i en array
    filnamn = Dir(rot_mapp & "\" & mapp & "\")
    i = 0
    Do While filnamn <> "" ' Loopa så länge det finns filnamn i mappen
        Debug.Print filnamn
        ' Lägg till filnamnet i arrayen
        ReDim Preserve filnamnsArray(i)
        filnamnsArray(i) = filnamn
        i = i + 1
        ' Läs nästa filnamn
        filnamn = Dir
    Loop
    
    i = 0
    j = 0
    Debug.Print "After while loop completed"
    
    ' Sök efter filnamn som börjar med "Balans" och lägg dem i en annan array
    For j = LBound(filnamnsArray) To UBound(filnamnsArray)
        If Left(filnamnsArray(j), 8) = "Huvudbok" Then
            ReDim Preserve huvudbokArray(i)
            huvudbokArray(i) = filnamnsArray(j)
            i = i + 1
        End If
    Next j
     Debug.Print "huvudbokArray filled"
    ' Skriv ut alla filnamn som börjar med "Huvudbok" i arrayen
     If Not IsEmpty(huvudbokArray) Then
        For i = LBound(huvudbokArray) To UBound(huvudbokArray)
            Debug.Print huvudbokArray(i)
        Next i
    Else
        Debug.Print "Inga resultatfiler hittades."
    End If

     
    i = 0
    j = 0
    Debug.Print "Before xlsarray filling"
    ' Sök efter filnamn som slutar med ".xlsx" och lägg dem i en annan array
    For j = LBound(huvudbokArray) To UBound(huvudbokArray)
        If Right(huvudbokArray(j), 5) = ".xlsx" Then
            ReDim Preserve xlsxArray(i)
            xlsxArray(i) = huvudbokArray(j)
            i = i + 1
        End If
    Next j
    Debug.Print "xlsarray filled"
    ' Skriv ut alla filnamn som slutar med ".xlsx" i arrayen
    For i = LBound(xlsxArray) To UBound(xlsxArray)
       Debug.Print xlsxArray(i)
     Next i
    


    ' Skriv ut alla filnamn som uppfyller villkoren i arrayen
    Dim parts() As String
    Dim hittad As Boolean
    Dim date_parts() As String
    hittad = False
    Dim file_start_date As Date
    Dim file_end_date As Date
    Dim rätt_huvudbok As String
    
    Dim year_start As Integer
    Dim month_start As Integer
    Dim day_start As Integer
    Dim year_end As Integer
    Dim month_end As Integer
    Dim day_end As Integer
    
    For i = LBound(xlsxArray) To UBound(xlsxArray)
        Debug.Print xlsxArray(i)
        rätt_huvudbok = xlsxArray(i)
    Next i
        
        
    ' Dim monthAbbreviation
    ' monthAbbreviation = Left(MonthName(month_start), 3)

    ' Debug.Print monthAbbreviation
  
    
    ' Skapa sökvägen till filen
    Dim path As String
    path = rot_mapp & "\" & mapp & "\" & rätt_huvudbok
    Debug.Print "Sökvägen till huvudboken är:"
    Debug.Print path
    
    
    ' Öppna resultarrapporten
    Dim huvudbokWorkbook As Workbook
    Set huvudbokWorkbook = Workbooks.Open(path)
    Dim huvudbokSheet As Worksheet
    Set huvudbokSheet = huvudbokWorkbook.Sheets(1) ' Antag att huvudboken är på första arket

   
    targetSheet.Activate
    
    
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    Debug.Print "Sista raden för resultatrapporten: " & lastRow
    
    ' Beräkna var du ska börja kopiera huvudboken
    Dim startRow As Long
    startRow = lastRow + 5 ' 5 rader under den sista raden i balansrapporten
    Debug.Print "Huvudboken skrivs in vid rad " & startRow
    
    
    ' Kopiera huvudboken till det angivna området
    huvudbokSheet.UsedRange.Copy targetSheet.Cells(startRow, 1) ' Börja i den första kolumnen på startRow
    
    ' Stäng resultatrapporten utan att spara ändringar
    huvudbokWorkbook.Close SaveChanges:=False
    
    targetSheet.Cells(3, 7).Value = startRow
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    targetSheet.Cells(3, 8).Value = lastRow
    targetSheet.Cells(3, 6).Value = "Huvudbok"
    

    Debug.Print "test"
End Sub

Sub LäsInVerifikationslista(ByVal targetSheet As Worksheet, ByVal rot_mapp As String, ByVal mapp As String, ByVal start_date As Date, ByVal end_date As Date)
    Dim filnamn As String
    Dim bladNamn As String
    Dim kolumnIndex As Integer
    Dim filnamnsArray() As String
    Dim verifikationslistaArray() As String
    Dim xlsxArray() As String
    Dim i As Integer
    Dim j As Integer

    
    
    
    ' Skriv ut sökvägen till mappen
    Debug.Print rot_mapp & "\" & mapp
    
    ' Återställ arrayen för att undvika eventuellt tidigare innehåll
    ReDim filnamnsArray(0)
    
    ' Läs in filnamnen som ligger i mappen och lägg filnamnen i en array
    filnamn = Dir(rot_mapp & "\" & mapp & "\")
    i = 0
    Do While filnamn <> "" ' Loopa så länge det finns filnamn i mappen
        Debug.Print filnamn
        ' Lägg till filnamnet i arrayen
        ReDim Preserve filnamnsArray(i)
        filnamnsArray(i) = filnamn
        i = i + 1
        ' Läs nästa filnamn
        filnamn = Dir
    Loop
    
    i = 0
    j = 0
    Debug.Print "After while loop completed"
    
    ' Sök efter filnamn som börjar med "Balans" och lägg dem i en annan array
    For j = LBound(filnamnsArray) To UBound(filnamnsArray)
        If Left(filnamnsArray(j), 12) = "Verifikation" Then
            ReDim Preserve verifikationslistaArray(i)
            verifikationslistaArray(i) = filnamnsArray(j)
            i = i + 1
        End If
    Next j
     Debug.Print "verifikationslistaArray filled"
    ' Skriv ut alla filnamn som börjar med "Huvudbok" i arrayen
     If Not IsEmpty(verifikationslistaArray) Then
        For i = LBound(verifikationslistaArray) To UBound(verifikationslistaArray)
            Debug.Print verifikationslistaArray(i)
        Next i
    Else
        Debug.Print "Inga verifikationslistafiler hittades."
    End If

     
    i = 0
    j = 0
    Debug.Print "Before xlsarray filling"
    ' Sök efter filnamn som slutar med ".xlsx" och lägg dem i en annan array
    For j = LBound(verifikationslistaArray) To UBound(verifikationslistaArray)
        If Right(verifikationslistaArray(j), 5) = ".xlsx" Then
            ReDim Preserve xlsxArray(i)
            xlsxArray(i) = verifikationslistaArray(j)
            i = i + 1
        End If
    Next j
    Debug.Print "xlsarray filled"
    ' Skriv ut alla filnamn som slutar med ".xlsx" i arrayen
    For i = LBound(xlsxArray) To UBound(xlsxArray)
       Debug.Print xlsxArray(i)
     Next i
    


    ' Skriv ut alla filnamn som uppfyller villkoren i arrayen
    Dim parts() As String
    Dim hittad As Boolean
    Dim date_parts() As String
    hittad = False
    Dim file_start_date As Date
    Dim file_end_date As Date
    Dim rätt_verifikationslista As String
    
    Dim year_start As Integer
    Dim month_start As Integer
    Dim day_start As Integer
    Dim year_end As Integer
    Dim month_end As Integer
    Dim day_end As Integer
    
    For i = LBound(xlsxArray) To UBound(xlsxArray)
        Debug.Print xlsxArray(i)
        rätt_verifikationslista = xlsxArray(i)
    Next i
        
        
    ' Dim monthAbbreviation
    ' monthAbbreviation = Left(MonthName(month_start), 3)

    ' Debug.Print monthAbbreviation
  
    
    ' Skapa sökvägen till filen
    Dim path As String
    path = rot_mapp & "\" & mapp & "\" & rätt_verifikationslista
    Debug.Print "Sökvägen till verifikationslistan är:"
    Debug.Print path
    
    
    ' Öppna resultarrapporten
    Dim verifikationslistaWorkbook As Workbook
    Set verifikationslistaWorkbook = Workbooks.Open(path)
    Dim verifikationslistaSheet As Worksheet
    Set verifikationslistaSheet = verifikationslistaWorkbook.Sheets(1) ' Antag att huvudboken är på första arket

   
    targetSheet.Activate
    
    
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    Debug.Print "Sista raden för huvudboken: " & lastRow
    
    ' Beräkna var du ska börja kopiera huvudboken
    Dim startRow As Long
    startRow = lastRow + 5 ' 5 rader under den sista raden i balansrapporten
    Debug.Print "verifikationslistan skrivs in vid rad " & startRow
    
    
    ' Kopiera verifikationslista till det angivna området
    verifikationslistaSheet.UsedRange.Copy targetSheet.Cells(startRow, 1) ' Börja i den första kolumnen på startRow
    
    ' Stäng resultatrapporten utan att spara ändringar
    verifikationslistaWorkbook.Close SaveChanges:=False
    
    targetSheet.Cells(4, 7).Value = startRow
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    targetSheet.Cells(4, 8).Value = lastRow
    targetSheet.Cells(4, 6).Value = "Verifikationslistan"
    

    Debug.Print "test"
End Sub
