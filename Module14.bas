Attribute VB_Name = "Module14"
Sub L�sInBalansrapport()
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
    
    ' Ange mappen d�r balansrapporten finns
    rot_mapp = "G:\Bokf�ring\Planering inf�r �rsbokslut"
    
    ' L�s in namnet p� den m�nadsmapp som balansrapporten ska h�mtas ifr�n
    mapp = ws.Cells(1, 1).Value
    
    ' Skriv ut s�kv�gen till mappen
    Debug.Print rot_mapp & "\" & mapp
    
    ' �terst�ll arrayen f�r att undvika eventuellt tidigare inneh�ll
    ReDim filnamnsArray(0)
    
    ' L�s in filnamnen som ligger i mappen och l�gg filnamnen i en array
    filnamn = Dir(rot_mapp & "\" & mapp & "\")
    i = 0
    Do While filnamn <> "" ' Loopa s� l�nge det finns filnamn i mappen
        Debug.Print filnamn
        ' L�gg till filnamnet i arrayen
        ReDim Preserve filnamnsArray(i)
        filnamnsArray(i) = filnamn
        i = i + 1
        ' L�s n�sta filnamn
        filnamn = Dir
    Loop
    
    i = 0
    j = 0
    Debug.Print "After while loop completed"
    
    ' S�k efter filnamn som b�rjar med "Balans" och l�gg dem i en annan array
    For j = LBound(filnamnsArray) To UBound(filnamnsArray)
        If Left(filnamnsArray(j), 6) = "Balans" Then
            ReDim Preserve balansArray(i)
            balansArray(i) = filnamnsArray(j)
            i = i + 1
        End If
    Next j
     Debug.Print "Balansarray filled"
    ' Skriv ut alla filnamn som b�rjar med "Balans" i arrayen
     For i = LBound(balansArray) To UBound(balansArray)
       Debug.Print balansArray(i)
     Next i
     
    i = 0
    j = 0
    Debug.Print "Before xlsarray filling"
    ' S�k efter filnamn som slutar med ".xlsx" och l�gg dem i en annan array
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
    Dim r�tt_balansrapport As String
    
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
        
        ' Kontrollera att det finns minst tv� delar efter uppdelning
        If UBound(parts) >= 1 Then
            ' Splitta den andra delen vid minustecknet
            date_parts = Split(parts(1), "-")
            Debug.Print date_parts(0)
            Debug.Print date_parts(1)
            
        End If
        ' Kontrollera att det finns exakt tv� delar efter uppdelning
        If UBound(date_parts) = 1 Then
            ' Extrahera start- och slutdatum fr�n filnamnet
            ' Dela upp start- och slutdatumet i �r, m�nad och dag
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
                ' S�tt flaggan och lagra indexet f�r det matchande filnamnet
                r�tt_balansrapport = xlsxArray(i)
                hittad = True
                Exit For ' Avbryt loopen n�r r�tt filnamn har hittats
            End If
        End If
    Next i
    
    If hittad Then
        ' Skriv ut det valda filnamnet
        Debug.Print "R�tt balansrapport: " & r�tt_balansrapport
    Else
        ' Filen hittas inte
        Debug.Print "R�tt balansrapport har inte hittats"
    End If
    
    Dim monthAbbreviation
    monthAbbreviation = Left(MonthName(month_start), 3)

    Debug.Print monthAbbreviation
    
    Dim targetSheet As Worksheet
    Dim sheetName As String
    sheetName = monthAbbreviation ' Anv�nd inneh�llet i monthAbbreviation som fliknamn

    ' Kontrollera om fliken med det aktuella namnet redan finns
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' Om fliken inte finns, skapa den
    If targetSheet Is Nothing Then
        Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        targetSheet.Name = sheetName
    End If
    
    
    
    ' Skapa s�kv�gen till filen
    Dim path As String
    path = rot_mapp & "\" & mapp & "\" & r�tt_balansrapport
    Debug.Print "S�kv�gen till balansrapporten �r:"
    Debug.Print path
    
    
    ' �ppna den andra filen
    Dim balansWorkbook As Workbook
    Set balansWorkbook = Workbooks.Open(path)

   
    targetSheet.Activate
    TaBortKolumner targetSheet

    ' Kopiera inneh�llet fr�n den enda fliken i balansWorkbook till targetSheet
    balansWorkbook.Sheets(1).UsedRange.Copy targetSheet.Range("A1")

    ' St�ng den andra filen utan att spara �ndringar
    balansWorkbook.Close SaveChanges:=False

    JusteraKolumnBredden targetSheet
    
    InfogaHeaders targetSheet
    
    targetSheet.Cells(1, 7).Value = file_start_date
    targetSheet.Cells(1, 8).Value = file_end_date
    
    L�sInResultatrapport targetSheet, rot_mapp, mapp, file_start_date, file_end_date
    L�sInHuvudbok targetSheet, rot_mapp, mapp, file_start_date, file_end_date
    L�sInVerifikationslista targetSheet, rot_mapp, mapp, file_start_date, file_end_date
    Debug.Print "test"
End Sub
Sub TaBortKolumner(ws As Worksheet)
    ' Ta bort kolumnerna A till J
   ws.Columns("A:J").Clear
End Sub


Sub JusteraKolumnBredden(ByVal targetSheet As Worksheet)
    Dim columnRange As Range
    Dim i As Integer
    
    ' Ange vilka kolumner du vill justera bredden f�r
    ' H�r antar jag att du vill justera alla kolumner fr�n A till Z
    Set columnRange = targetSheet.Range("A:K")
    
    ' Ange den �nskade bredden i tecken
    Dim desiredWidth As Integer
    desiredWidth = 15 ' Du kan �ndra detta v�rde till det du anser passar b�st
    
    ' Justera bredden f�r varje kolumn i den angivna kolumnr�ckan
    For i = 1 To columnRange.Columns.Count
        columnRange.Columns(i).ColumnWidth = desiredWidth
    Next i
End Sub

Sub InfogaHeaders(ByVal targetSheet As Worksheet)
    Dim startRow As Long
    Dim lastRow As Long
    Dim headerRange As Range
    
    ' Ange startraden f�r att s�ka efter kriterierna
    startRow = 1
    
    ' Hitta den sista raden med data i f�rsta kolumnen
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    
    ' S�k igenom kolumn A f�r att hitta de specificerade kriterierna och infoga kolumnheaders
    For i = startRow To lastRow
        If targetSheet.Cells(i, 1).Value = "Materiella anl�ggningstillg�ngar" Or _
           targetSheet.Cells(i, 1).Value = "Kortfristiga fordringar" Or _
           targetSheet.Cells(i, 1).Value = "EGET KAPITAL, AVS�TTNINGAR OCH SKULDER" Or _
           targetSheet.Cells(i, 1).Value = "L�ngfristiga skulder" Or _
           targetSheet.Cells(i, 1).Value = "Kortfristiga skulder" Then
           
           ' Infoga kolumnheaders med fetstil fr�n kolumn C
           Set headerRange = targetSheet.Range("C" & i)
           headerRange.Resize(1, 17).Value = Array("Ing balans", "Ing saldo", "Period", _
                                                   "Utg balans", "Period ber�knad", _
                                                   "Utg balans ber�knad", "�verensst�mmer", _
                                                   "Ber�kningsunderlag", "1", "2", "3", "4", "5", "6", "7", "IB koll", "Saldo koll")
           headerRange.Resize(1, 17).Font.Bold = True
           
           
        End If
    Next i
End Sub

Sub L�sInResultatrapport(ByVal targetSheet As Worksheet, ByVal rot_mapp As String, ByVal mapp As String, ByVal start_date As Date, ByVal end_date As Date)
    Dim filnamn As String
    Dim bladNamn As String
    Dim kolumnIndex As Integer
    Dim filnamnsArray() As String
    Dim resultatArray() As String
    Dim xlsxArray() As String
    Dim i As Integer
    Dim j As Integer

    
    
    
    ' Skriv ut s�kv�gen till mappen
    Debug.Print rot_mapp & "\" & mapp
    
    ' �terst�ll arrayen f�r att undvika eventuellt tidigare inneh�ll
    ReDim filnamnsArray(0)
    
    ' L�s in filnamnen som ligger i mappen och l�gg filnamnen i en array
    filnamn = Dir(rot_mapp & "\" & mapp & "\")
    i = 0
    Do While filnamn <> "" ' Loopa s� l�nge det finns filnamn i mappen
        Debug.Print filnamn
        ' L�gg till filnamnet i arrayen
        ReDim Preserve filnamnsArray(i)
        filnamnsArray(i) = filnamn
        i = i + 1
        ' L�s n�sta filnamn
        filnamn = Dir
    Loop
    
    i = 0
    j = 0
    Debug.Print "After while loop completed"
    
    ' S�k efter filnamn som b�rjar med "Balans" och l�gg dem i en annan array
    For j = LBound(filnamnsArray) To UBound(filnamnsArray)
        If Left(filnamnsArray(j), 8) = "Resultat" Then
            ReDim Preserve resultatArray(i)
            resultatArray(i) = filnamnsArray(j)
            i = i + 1
        End If
    Next j
     Debug.Print "ResultatArray filled"
    ' Skriv ut alla filnamn som b�rjar med "Resultat" i arrayen
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
    ' S�k efter filnamn som slutar med ".xlsx" och l�gg dem i en annan array
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
    Dim r�tt_resultatrapport As String
    
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
        
        ' Kontrollera att det finns minst tv� delar efter uppdelning
        If UBound(parts) >= 1 Then
            ' Splitta den andra delen vid minustecknet
            date_parts = Split(parts(1), "-")
            Debug.Print date_parts(0)
            Debug.Print date_parts(1)
            
        End If
        ' Kontrollera att det finns exakt tv� delar efter uppdelning
        If UBound(date_parts) = 1 Then
            ' Extrahera start- och slutdatum fr�n filnamnet
            ' Dela upp start- och slutdatumet i �r, m�nad och dag
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
                ' S�tt flaggan och lagra indexet f�r det matchande filnamnet
                r�tt_resultatrapport = xlsxArray(i)
                hittad = True
                Exit For ' Avbryt loopen n�r r�tt filnamn har hittats
            End If
        End If
    Next i
    
    If hittad Then
        ' Skriv ut det valda filnamnet
        Debug.Print "R�tt resultatrapport: " & r�tt_resultatrapport
    Else
        ' Filen hittas inte
        Debug.Print "R�tt resultatrapport har inte hittats"
    End If
    
    Dim monthAbbreviation
    monthAbbreviation = Left(MonthName(month_start), 3)

    Debug.Print monthAbbreviation
  
    
    ' Skapa s�kv�gen till filen
    Dim path As String
    path = rot_mapp & "\" & mapp & "\" & r�tt_resultatrapport
    Debug.Print "S�kv�gen till resultatrapporten �r:"
    Debug.Print path
    
    
    ' �ppna resultarrapporten
    Dim resultatWorkbook As Workbook
    Set resultatWorkbook = Workbooks.Open(path)
    Dim resultatSheet As Worksheet
    Set resultatSheet = resultatWorkbook.Sheets(1) ' Antag att resultatrapporten �r p� f�rsta arket

   
    targetSheet.Activate
    
    
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    
    ' Ber�kna var du ska b�rja kopiera resultatrapporten
    Dim startRow As Long
    startRow = lastRow + 5 ' 5 rader under den sista raden i balansrapporten
    
    
    ' Kopiera resultatrapporten till det angivna omr�det
    resultatSheet.UsedRange.Copy targetSheet.Cells(startRow, 1) ' B�rja i den f�rsta kolumnen p� startRow
    
    ' St�ng resultatrapporten utan att spara �ndringar
    resultatWorkbook.Close SaveChanges:=False
    
    targetSheet.Cells(2, 7).Value = startRow
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    targetSheet.Cells(2, 8).Value = lastRow
    targetSheet.Cells(2, 6).Value = "Resultatrapport"
    
    

    Debug.Print "test"
End Sub


Sub L�sInHuvudbok(ByVal targetSheet As Worksheet, ByVal rot_mapp As String, ByVal mapp As String, ByVal start_date As Date, ByVal end_date As Date)
    Dim filnamn As String
    Dim bladNamn As String
    Dim kolumnIndex As Integer
    Dim filnamnsArray() As String
    Dim huvudbokArray() As String
    Dim xlsxArray() As String
    Dim i As Integer
    Dim j As Integer

    
    
    
    ' Skriv ut s�kv�gen till mappen
    Debug.Print rot_mapp & "\" & mapp
    
    ' �terst�ll arrayen f�r att undvika eventuellt tidigare inneh�ll
    ReDim filnamnsArray(0)
    
    ' L�s in filnamnen som ligger i mappen och l�gg filnamnen i en array
    filnamn = Dir(rot_mapp & "\" & mapp & "\")
    i = 0
    Do While filnamn <> "" ' Loopa s� l�nge det finns filnamn i mappen
        Debug.Print filnamn
        ' L�gg till filnamnet i arrayen
        ReDim Preserve filnamnsArray(i)
        filnamnsArray(i) = filnamn
        i = i + 1
        ' L�s n�sta filnamn
        filnamn = Dir
    Loop
    
    i = 0
    j = 0
    Debug.Print "After while loop completed"
    
    ' S�k efter filnamn som b�rjar med "Balans" och l�gg dem i en annan array
    For j = LBound(filnamnsArray) To UBound(filnamnsArray)
        If Left(filnamnsArray(j), 8) = "Huvudbok" Then
            ReDim Preserve huvudbokArray(i)
            huvudbokArray(i) = filnamnsArray(j)
            i = i + 1
        End If
    Next j
     Debug.Print "huvudbokArray filled"
    ' Skriv ut alla filnamn som b�rjar med "Huvudbok" i arrayen
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
    ' S�k efter filnamn som slutar med ".xlsx" och l�gg dem i en annan array
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
    Dim r�tt_huvudbok As String
    
    Dim year_start As Integer
    Dim month_start As Integer
    Dim day_start As Integer
    Dim year_end As Integer
    Dim month_end As Integer
    Dim day_end As Integer
    
    For i = LBound(xlsxArray) To UBound(xlsxArray)
        Debug.Print xlsxArray(i)
        r�tt_huvudbok = xlsxArray(i)
    Next i
        
        
    ' Dim monthAbbreviation
    ' monthAbbreviation = Left(MonthName(month_start), 3)

    ' Debug.Print monthAbbreviation
  
    
    ' Skapa s�kv�gen till filen
    Dim path As String
    path = rot_mapp & "\" & mapp & "\" & r�tt_huvudbok
    Debug.Print "S�kv�gen till huvudboken �r:"
    Debug.Print path
    
    
    ' �ppna resultarrapporten
    Dim huvudbokWorkbook As Workbook
    Set huvudbokWorkbook = Workbooks.Open(path)
    Dim huvudbokSheet As Worksheet
    Set huvudbokSheet = huvudbokWorkbook.Sheets(1) ' Antag att huvudboken �r p� f�rsta arket

   
    targetSheet.Activate
    
    
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    Debug.Print "Sista raden f�r resultatrapporten: " & lastRow
    
    ' Ber�kna var du ska b�rja kopiera huvudboken
    Dim startRow As Long
    startRow = lastRow + 5 ' 5 rader under den sista raden i balansrapporten
    Debug.Print "Huvudboken skrivs in vid rad " & startRow
    
    
    ' Kopiera huvudboken till det angivna omr�det
    huvudbokSheet.UsedRange.Copy targetSheet.Cells(startRow, 1) ' B�rja i den f�rsta kolumnen p� startRow
    
    ' St�ng resultatrapporten utan att spara �ndringar
    huvudbokWorkbook.Close SaveChanges:=False
    
    targetSheet.Cells(3, 7).Value = startRow
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    targetSheet.Cells(3, 8).Value = lastRow
    targetSheet.Cells(3, 6).Value = "Huvudbok"
    

    Debug.Print "test"
End Sub

Sub L�sInVerifikationslista(ByVal targetSheet As Worksheet, ByVal rot_mapp As String, ByVal mapp As String, ByVal start_date As Date, ByVal end_date As Date)
    Dim filnamn As String
    Dim bladNamn As String
    Dim kolumnIndex As Integer
    Dim filnamnsArray() As String
    Dim verifikationslistaArray() As String
    Dim xlsxArray() As String
    Dim i As Integer
    Dim j As Integer

    
    
    
    ' Skriv ut s�kv�gen till mappen
    Debug.Print rot_mapp & "\" & mapp
    
    ' �terst�ll arrayen f�r att undvika eventuellt tidigare inneh�ll
    ReDim filnamnsArray(0)
    
    ' L�s in filnamnen som ligger i mappen och l�gg filnamnen i en array
    filnamn = Dir(rot_mapp & "\" & mapp & "\")
    i = 0
    Do While filnamn <> "" ' Loopa s� l�nge det finns filnamn i mappen
        Debug.Print filnamn
        ' L�gg till filnamnet i arrayen
        ReDim Preserve filnamnsArray(i)
        filnamnsArray(i) = filnamn
        i = i + 1
        ' L�s n�sta filnamn
        filnamn = Dir
    Loop
    
    i = 0
    j = 0
    Debug.Print "After while loop completed"
    
    ' S�k efter filnamn som b�rjar med "Balans" och l�gg dem i en annan array
    For j = LBound(filnamnsArray) To UBound(filnamnsArray)
        If Left(filnamnsArray(j), 12) = "Verifikation" Then
            ReDim Preserve verifikationslistaArray(i)
            verifikationslistaArray(i) = filnamnsArray(j)
            i = i + 1
        End If
    Next j
     Debug.Print "verifikationslistaArray filled"
    ' Skriv ut alla filnamn som b�rjar med "Huvudbok" i arrayen
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
    ' S�k efter filnamn som slutar med ".xlsx" och l�gg dem i en annan array
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
    Dim r�tt_verifikationslista As String
    
    Dim year_start As Integer
    Dim month_start As Integer
    Dim day_start As Integer
    Dim year_end As Integer
    Dim month_end As Integer
    Dim day_end As Integer
    
    For i = LBound(xlsxArray) To UBound(xlsxArray)
        Debug.Print xlsxArray(i)
        r�tt_verifikationslista = xlsxArray(i)
    Next i
        
        
    ' Dim monthAbbreviation
    ' monthAbbreviation = Left(MonthName(month_start), 3)

    ' Debug.Print monthAbbreviation
  
    
    ' Skapa s�kv�gen till filen
    Dim path As String
    path = rot_mapp & "\" & mapp & "\" & r�tt_verifikationslista
    Debug.Print "S�kv�gen till verifikationslistan �r:"
    Debug.Print path
    
    
    ' �ppna resultarrapporten
    Dim verifikationslistaWorkbook As Workbook
    Set verifikationslistaWorkbook = Workbooks.Open(path)
    Dim verifikationslistaSheet As Worksheet
    Set verifikationslistaSheet = verifikationslistaWorkbook.Sheets(1) ' Antag att huvudboken �r p� f�rsta arket

   
    targetSheet.Activate
    
    
    Dim lastRow As Long
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    Debug.Print "Sista raden f�r huvudboken: " & lastRow
    
    ' Ber�kna var du ska b�rja kopiera huvudboken
    Dim startRow As Long
    startRow = lastRow + 5 ' 5 rader under den sista raden i balansrapporten
    Debug.Print "verifikationslistan skrivs in vid rad " & startRow
    
    
    ' Kopiera verifikationslista till det angivna omr�det
    verifikationslistaSheet.UsedRange.Copy targetSheet.Cells(startRow, 1) ' B�rja i den f�rsta kolumnen p� startRow
    
    ' St�ng resultatrapporten utan att spara �ndringar
    verifikationslistaWorkbook.Close SaveChanges:=False
    
    targetSheet.Cells(4, 7).Value = startRow
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    targetSheet.Cells(4, 8).Value = lastRow
    targetSheet.Cells(4, 6).Value = "Verifikationslistan"
    

    Debug.Print "test"
End Sub
