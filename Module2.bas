Attribute VB_Name = "Module2"
Sub TestaIngSaldo()
    Dim ws As Worksheet
    Dim start_date As Date
    Dim end_date As Date
    Dim currentMonth As Integer
    Dim monthAbbreviation As String
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim dataRange As Range
    Dim data() As Variant
    
    
    Debug.Print "Inne i TestaIngSaldo()"
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
    
    If monthAbbreviation = "Jan" Then
    ' första månaden i perioden
    'ingående saldo måste vara IB
    ' jämför ingående saldo med IB
    ' och om lika skriv ingående saldot på rad kolumn S
   
    
             ' Hitta den sista raden med data i kolumn A
        lastRow = targetSheet.Cells(2, 7).Value - 1
        Debug.Print "lastRow är: " & lastRow
        
         'Aktivera fliken
        targetSheet.Activate
        
        ' Hitta den sista raden med data i kolumn A
        lastRow = targetSheet.Cells(2, 7).Value - 1
        Debug.Print "lastRow är: " & lastRow
        
        ' Ange området med data
        Set dataRange = targetSheet.Range("C1:D" & lastRow)
        ReDim Preserve data(1 To lastRow, 1 To 2)
        
        ' Läs in datan till en tvådimensionell array
        data = dataRange.Value
        
        ' Skriv ut innehållet i variabeln data för att verifiera om den är tom eller inte
        Debug.Print "Innehållet i variabeln data:"
        For i = LBound(data, 1) To UBound(data, 1)
            Debug.Print "Rad " & i & ":" & data(i, 1) & " " & data(i, 2)
        Next i
        
        
                ' Loopa igenom varje rad i datan
        For i = 1 To UBound(data, 1)
            ' Kontrollera om både IB och ingående saldo är numeriska värden och inte tomma
            If IsNumeric(data(i, 1)) And IsNumeric(data(i, 2)) And data(i, 1) <> "" And data(i, 2) <> "" Then
                ' Jämför IB och ingående saldo för varje rad
                If data(i, 1) = data(i, 2) Then
                    ' Om IB och ingående saldo är lika, skriv IB-värdet till kolumn S på samma rad
                    targetSheet.Cells(i, 19).Value = data(i, 1)
                End If
            End If
        Next i

    Else
        MsgBox "Det är inte Januari och nen mer komplex verifiering av Ingående saldo måste göras, t.ex."
    End If
    
End Sub
