Attribute VB_Name = "Module1"
Sub Stock_Analysis()
Dim stockOpen As Double
Dim stockClose As Double
Dim yearlyChange As Double
Dim ticker As String
Dim outputRow As Integer
Dim volume As LongLong
Dim percentChange As Double

For Each ws In Worksheets
    outputRow = 2

' Finding yearly change for each ticker
    stockOpen = ws.Cells(2, 3).Value
    stockClose = ws.Cells(2, 6).Value
    yearlyChange = 0
    volume = ws.Cells(2, 7).Value
         For i = 2 To (ws.Range("A2").End(xlDown).Row)

            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                stockClose = ws.Cells(i + 1, 6).Value
'Finding total volume is the sum of all volumes
                volume = volume + ws.Cells(i + 1, 7).Value
                ticker = ws.Cells(i, 1).Value
            Else
'Finding yearly change is the final close - initial open
                yearlyChange = stockClose - stockOpen
' Calculating percent change if opened at 0
                If stockOpen = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / stockOpen
                End If
'Putting data for individual stock in the table
                ws.Cells(outputRow, 10).Value = ticker
                ws.Cells(outputRow, 11).Value = yearlyChange
                ws.Cells(outputRow, 12).Value = percentChange
                ws.Cells(outputRow, 13).Value = volume
'Setting initial stock values for the next ticker
                stockOpen = ws.Cells(i + 1, 3).Value
                stockClose = ws.Cells(i + 1, 6).Value
                volume = ws.Cells(i + 1, 7).Value
                yearlyChange = 0
                outputRow = outputRow + 1
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
