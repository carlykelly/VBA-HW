Attribute VB_Name = "Module1"

Sub stocks()

Dim stockOpen As Double
Dim stockClose As Double
Dim yearlyChange As Double
Dim ticker As String
Dim a As Integer
Dim volume As LongLong
Dim percentChange As Double

For Each ws In Worksheets
a = 2

' Finding yearly change for each ticker
    stockOpen = ws.Cells(2, 3).Value
    stockClose = ws.Cells(2, 6).Value
    yearlyChange = 0
    volume = ws.Cells(2, 7).Value
         For i = 2 To (ws.Range("A2").End(xlDown).Row)

            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                stockClose = ws.Cells(i + 1, 6).Value
'Assuming total volume is the sum of all volumes
                volume = volume + ws.Cells(i + 1, 7).Value
                ticker = ws.Cells(i, 1).Value
            Else
'Assuming yearly change is the final close - initial open
                yearlyChange = stockClose - stockOpen
' Calculating percent change if opened at 0
                If stockOpen = 0 Then
                    ws.Cells(a, 12) = 0
                Else
                percentChange = yearlyChange / stockOpen
                ws.Cells(a, 12).Value = percentChange
                End If
                ws.Cells(a, 11).Value = yearlyChange
                ws.Cells(a, 10).Value = ticker
                ws.Cells(a, 13).Value = volume
                stockOpen = ws.Cells(i + 1, 3).Value
                stockClose = ws.Cells(i + 1, 6).Value
                volume = ws.Cells(i + 1, 7).Value
                yearlyChange = 0
                a = a + 1
            End If
            
        Next i
'Labeling table created
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 13).Value = "Volume"
ws.Cells(1, 12).Value = "Percent Change"
' Conditional Formatting
    For i = 2 To (ws.Range("K2").End(xlDown).Row)
        If ws.Cells(i, 11).Value < 0 Then
            ws.Cells(i, 11).Interior.Color = RGB(255, 0, 0)
        Else
            ws.Cells(i, 11).Interior.Color = RGB(0, 255, 0)
        End If
    Next i
        
   For i = 2 To (ws.Range("L2").End(xlDown).Row)
    ws.Cells(i, 12).NumberFormat = "0.00%"
    Next i
Next ws
End Sub
