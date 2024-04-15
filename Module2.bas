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
    
    If monthAbbreviation = "Jan" Then
    ' f�rsta m�naden i perioden
    'ing�ende saldo m�ste vara IB
    ' j�mf�r ing�ende saldo med IB
    ' och om lika skriv ing�ende saldot p� rad kolumn S
   
    
             ' Hitta den sista raden med data i kolumn A
        lastRow = targetSheet.Cells(2, 7).Value - 1
        Debug.Print "lastRow �r: " & lastRow
        
         'Aktivera fliken
        targetSheet.Activate
        
        ' Hitta den sista raden med data i kolumn A
        lastRow = targetSheet.Cells(2, 7).Value - 1
        Debug.Print "lastRow �r: " & lastRow
        
        ' Ange omr�det med data
        Set dataRange = targetSheet.Range("C1:D" & lastRow)
        ReDim Preserve data(1 To lastRow, 1 To 2)
        
        ' L�s in datan till en tv�dimensionell array
        data = dataRange.Value
        
        ' Skriv ut inneh�llet i variabeln data f�r att verifiera om den �r tom eller inte
        Debug.Print "Inneh�llet i variabeln data:"
        For i = LBound(data, 1) To UBound(data, 1)
            Debug.Print "Rad " & i & ":" & data(i, 1) & " " & data(i, 2)
        Next i
        
        
                ' Loopa igenom varje rad i datan
        For i = 1 To UBound(data, 1)
            ' Kontrollera om b�de IB och ing�ende saldo �r numeriska v�rden och inte tomma
            If IsNumeric(data(i, 1)) And IsNumeric(data(i, 2)) And data(i, 1) <> "" And data(i, 2) <> "" Then
                ' J�mf�r IB och ing�ende saldo f�r varje rad
                If data(i, 1) = data(i, 2) Then
                    ' Om IB och ing�ende saldo �r lika, skriv IB-v�rdet till kolumn S p� samma rad
                    targetSheet.Cells(i, 19).Value = data(i, 1)
                End If
            End If
        Next i

    Else
        MsgBox "Det �r inte Januari och nen mer komplex verifiering av Ing�ende saldo m�ste g�ras, t.ex."
    End If
    
End Sub
