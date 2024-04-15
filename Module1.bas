Attribute VB_Name = "Module1"
Sub TestaIB()
    Dim ws As Worksheet
    Dim start_date As Date
    Dim end_date As Date
    Dim currentMonth As Integer
    Dim monthAbbreviation As String
    Dim targetSheet As Worksheet
    Dim filteredData() As Variant ' Dynamiskt tv�dimensionellt array f�r filtrerad data
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim IbSheet As Worksheet
    Dim ibData() As Variant
    Dim ibRowsCount As Long
    
    
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
    End If
    
     ' Hitta den sista raden med data i kolumn A
    lastRow = targetSheet.Cells(2, 7).Value - 1
    Debug.Print "lastRow �r: " & lastRow
    
    ' Redimensionera den dynamiska arrayen till antalet rader i dataomr�det
    ' ReDim Preserve filteredData(lastRow)
    ' ReDim filteredData(lastRow, 3)
    ReDim filteredData(1 To lastRow, 1 To 3)

    
    ' Dim upperBound As Long
    ' upperBound = UBound(filteredData)
    ' Debug.Print "�vre gr�ns f�r filteredData: " & upperBound

    
    ' Anropa funktionen ReadAndFilterData f�r att filtrera data p� m�nadsfliken
    filteredData = ReadAndFilterData(targetSheet)
    
     ' Skriv ut inneh�llet i variabeln filteredData f�r att verifiera om den �r tom eller inte
    Debug.Print "Inneh�llet i variabeln filteredData:"
    For i = LBound(filteredData, 1) To UBound(filteredData, 1)
        Debug.Print "Rad " & i & ":" & "  (1.) " & filteredData(i, 1) & "  (2.) " & filteredData(i, 2) & " rad nummer: " & filteredData(i, 3)
    Next i
    
    ' Hitta fliken f�r de ing�ende balanserna
    Set IbSheet = ThisWorkbook.Sheets("IB")
    
    'Antal rader i IB -fliken
    ibRowsCount = IbSheet.Cells(IbSheet.Rows.Count, "A").End(xlUp).Row + 1
    Debug.Print "ibRowsCount: " & ibRowsCount
    
    ' Dimensionera arrayen
    ReDim ibData(1 To ibRowsCount, 1 To 3)
    Debug.Print "UBound(ibData, 1): " & UBound(ibData, 1)
    
    ' L�s in hela datan fr�n fliken "IB"
    ibData = IbSheet.Range("A1:C" & ibRowsCount)
    
     ' Skriv ut inneh�llet i variabeln ibData f�r att verifiera om den �r tom eller inte
    Debug.Print "Inneh�llet i variabeln ibData:"
    For i = LBound(ibData, 1) To UBound(ibData, 1)
        Debug.Print "Rad " & i & ":" & "  (1.) " & ibData(i, 1) & "    " & ibData(i, 2) & "   " & ibData(i, 3)
    Next i
    
    ' Iterera �ver varje rad i filteredData
    For i = LBound(filteredData, 1) To UBound(filteredData, 1)
        ' Kontrollera om v�rdet i f�rsta kolumnen �r numeriskt
        If IsNumeric(filteredData(i, 1)) Then
            ' S�k motsvarande kontonummer i kolumn A p� fliken "IB"
            For j = LBound(ibData, 1) To UBound(ibData, 1)
                If ibData(j, 1) = filteredData(i, 1) Then
                    ' Om kontonumret matchar, l�s in IB-v�rdet fr�n kolumn C
                    ibValue = ibData(j, 3)
                    ' J�mf�r IB-v�rdet med andra kolumnen i filteredData
                    If ibValue = filteredData(i, 2) Then
                        ' Om v�rdena �r lika, skriv IB-v�rdet till kolumn R p� targetSheet
                        targetSheet.Cells(filteredData(i, 3), 18).Value = ibValue ' Kolumn R = 18
                        ' Skapa en hyperl�nk fr�n balansrapportens IB-v�rde till fliken "IB" och motsvarande konto
                        Debug.Print "Anchor cell: " & targetSheet.Cells(filteredData(i, 3), 18).Address
                        Debug.Print "IbSheet name: " & IbSheet.Name
                        Debug.Print "SubAddress: '" & IbSheet.Name & "'!D" & j
                        Debug.Print "TextToDisplay: " & ibValue
                        targetSheet.Select
                        targetSheet.Cells(filteredData(i, 3), 18).Select

                        targetSheet.Hyperlinks.Add Anchor:=targetSheet.Cells(filteredData(i, 3), 18), Address:="", SubAddress:="" & IbSheet.Name & "!D" & j, TextToDisplay:=ibValue
                        Exit For ' G� till n�sta rad i filteredData
                    End If
                End If
            Next j
        End If
    Next i
    
     
    
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
    Debug.Print "lastRow �r: " & lastRow
    
    ' Ange omr�det med data
    Set dataRange = targetSheet.Range("A1:C" & lastRow)
    ReDim Preserve data(1 To lastRow, 1 To 3)
    
    ' L�s in datan till en tv�dimensionell array
    data = dataRange.Value
    
    ' Skriv ut inneh�llet i variabeln data f�r att verifiera om den �r tom eller inte
    Debug.Print "Inneh�llet i variabeln data:"
    For i = LBound(data, 1) To UBound(data, 1)
        Debug.Print "Rad " & i & ":" & data(i, 1) & " " & data(i, 2) & " " & data(i, 3)
    Next i
    
    j = 1
    ' Skapa den slutgiltiga filteredData-arrayn med r�tt storlek
    ReDim filteredData(1 To lastRow, 1 To 3)

    
    ' Loopa igenom varje rad i datan
    For i = 1 To UBound(data, 1)
        ' Kontrollera om det finns en siffra i A-delen av raden
        If Not IsEmpty(data(i, 1)) And data(i, 1) Like "####" Then
             filteredData(j, 1) = data(i, 1)
             filteredData(j, 2) = data(i, 3)
             filteredData(j, 3) = i
             j = j + 1
        End If
    Next i
       
    ReadAndFilterData = filteredData
    Debug.Print "Returnerar h�r"

End Function
