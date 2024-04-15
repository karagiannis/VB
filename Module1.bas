Attribute VB_Name = "Module1"
Sub TestaIB()
    Dim ws As Worksheet
    Dim start_date As Date
    Dim end_date As Date
    Dim currentMonth As Integer
    Dim monthAbbreviation As String
    Dim targetSheet As Worksheet
    Dim filteredData() As Variant ' Dynamiskt tvådimensionellt array för filtrerad data
    Dim lastRow As Long
    
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
    End If
    
     ' Hitta den sista raden med data i kolumn A
    lastRow = targetSheet.Cells(2, 7).Value - 1
    Debug.Print "lastRow är: " & lastRow
    
    ' Redimensionera den dynamiska arrayen till antalet rader i dataområdet
    ' ReDim Preserve filteredData(lastRow)
    ' ReDim filteredData(lastRow, 3)
    ReDim filteredData(1 To lastRow, 1 To 2)

    
    ' Dim upperBound As Long
    ' upperBound = UBound(filteredData)
    ' Debug.Print "Övre gräns för filteredData: " & upperBound

    
    ' Anropa funktionen ReadAndFilterData för att filtrera data på månadsfliken
    filteredData = ReadAndFilterData(targetSheet)
    
    Debug.Print filteredData(1, 1)
    
End Sub


Function ReadAndFilterData(ByVal targetSheet As Worksheet) As Variant


    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim data() As Variant
    Dim filteredData() As Variant
    Dim i As Long, j As Long
    Dim hasNumber As Boolean

    
    'Aktivera fliken
    targetSheet.Activate
    
    ' Hitta den sista raden med data i kolumn A
    lastRow = targetSheet.Cells(2, 7).Value - 1
    Debug.Print "lastRow är: " & lastRow
    
    ' Ange området med data
    Set dataRange = targetSheet.Range("A1:C" & lastRow)
    ReDim Preserve data(1 To lastRow, 1 To 3)
    
    ' Läs in datan till en tvådimensionell array
    data = dataRange.Value
    
    ' Skriv ut innehållet i variabeln data för att verifiera om den är tom eller inte
    Debug.Print "Innehållet i variabeln data:"
    For i = LBound(data, 1) To UBound(data, 1)
        Debug.Print "Rad " & i & ":" & data(i, 1) & " " & data(i, 2) & " " & data(i, 3)
    Next i
    
    j = 1
    ' Skapa den slutgiltiga filteredData-arrayn med rätt storlek
    ReDim filteredData(1 To lastRow, 1 To 2)

    
    ' Loopa igenom varje rad i datan
    For i = 1 To UBound(data, 1)
        ' Kontrollera om det finns en siffra i A-delen av raden
        If IsNumeric(data(i, 1)) Then
             filteredData(j, 1) = data(i, 1)
             filteredData(j, 2) = data(i, 3)
             j = j + 1
        End If
    Next i
    
    
    ' Skriv ut innehållet i variabeln filteredData för att verifiera om den är tom eller inte
    Debug.Print "Innehållet i variabeln filteredData:"
    For i = LBound(filteredData, 1) To UBound(filteredData, 1)
        Debug.Print "Rad " & i & ":" & "  (1.) " & filteredData(i, 1) & "  (2.) " & filteredData(i, 2)
    Next i
        
    
    
    ReadAndFilterData = filteredData
    Debug.Print "Returnerar här"

End Function
