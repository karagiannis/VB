Attribute VB_Name = "Module1"
Sub TestaIB()
    Dim ws As Worksheet
    Dim start_date As Date
    Dim end_date As Date
    Dim currentMonth As Integer
    Dim monthAbbreviation As String
    Dim targetSheet As Worksheet
    Dim filteredData() As Variant ' Dynamiskt tvådimensionellt array för filtrerad data
    
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
    
    ' Redimensionera den dynamiska arrayen till antalet rader i dataområdet
    ReDim Preserve filteredData(lastRow)
    
    Dim upperBound As Long
    upperBound = UBound(filteredData)
    Debug.Print "Övre gräns för filteredData: " & upperBound

    
    ' Anropa funktionen ReadAndFilterData för att filtrera data på månadsfliken
    filteredData(1) = ReadAndFilterData(targetSheet)
    
    Debug.Print filteredData(1)
    
End Sub


Function ReadAndFilterData(ByVal targetSheet As Worksheet) As Variant


    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim data As Variant
    Dim filteredData As Variant
    Dim i As Long, j As Long
    Dim hasNumber As Boolean
    
    'Aktivera fliken
    targetSheet.Activate
    
    ' Hitta den sista raden med data i kolumn A
    lastRow = targetSheet.Cells(2, 7).Value - 1
    Debug.Print "lastRow är: " & lastRow
    
    ' Ange området med data
    Set dataRange = targetSheet.Range("A1:C" & lastRow)
    
    Dim result As Variant
    result = "Test"
    ReadAndFilterData = result

End Function
