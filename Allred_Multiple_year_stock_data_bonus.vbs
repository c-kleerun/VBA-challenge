VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockmarket_a()

Dim ws As Worksheet

Dim ticker As String

Dim yearlyChange As Double
yearlyChange = 0

Dim percentChange As Double
percentChange = 0

Dim totalVolume As Double

Dim summary_table_row As Integer

Dim openvalue As Double

For Each ws In Worksheets
WorksheetName = ws.Name
   MsgBox WorksheetName
    summary_table_row = 2
    totalVolume = 0
    openvalue = 2
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        yearlyChange = ws.Cells(i, 6).Value - ws.Cells(openvalue, 3).Value
    If ws.Cells(openvalue, 3).Value = 0 Then
        percentChange = 0
    Else
        percentChange = (yearlyChange / ws.Cells(openvalue, 3).Value)
    End If
            
        ws.Range("I" & summary_table_row).Value = ticker
        ws.Range("L" & summary_table_row).Value = totalVolume
        ws.Range("J" & summary_table_row).Value = yearlyChange
        ws.Range("K" & summary_table_row).Value = percentChange
        
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Volume")
        summary_table_row = summary_table_row + 1
        openvalue = i + 1

    Else
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    End If
Next i

For j = 2 To lastrow
    If ws.Cells(j, 10).Value < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
    ElseIf ws.Cells(j, 10).Value >= 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    End If

        ws.Cells(j, 11).Style = "Percent"
                
Next j

Dim max_volume As Double
Dim max_ticker As String
max_volume = 0

For x = 2 To summary_table_row
    If ws.Cells(x, 12).Value > max_volume Then
        max_volume = ws.Cells(x, 12)
        max_ticker = ws.Cells(x, 9)
    End If
Next x

Dim increase_percent As Double
Dim increase_ticker As String
increase_percent = 0
Dim decrease_percent As Double
Dim decrease_ticker As String
decrease_percent = 0

For y = 2 To summary_table_row
    If ws.Cells(y, 11).Value > increase_percent Then
        increase_percent = ws.Cells(y, 11)
        increase_ticker = ws.Cells(y, 9)
    End If
    
    If ws.Cells(y, 11).Value < decrease_percent Then
        decrease_percent = ws.Cells(y, 11)
        decrease_ticker = ws.Cells(y, 9)
    End If
    
Next y

ws.Range("O1:P1") = Array("Ticker", "Value")

ws.Cells(2, 14).Value = "Greatest Percent Increase"
ws.Cells(2, 16).Value = increase_percent
ws.Cells(2, 15).Value = increase_ticker
ws.Cells(2, 16).Style = "Percent"

ws.Cells(3, 14).Value = "Greatest Percent Decrease"
ws.Cells(3, 16).Value = decrease_percent
ws.Cells(3, 15).Value = decrease_ticker
ws.Cells(3, 16).Style = "Percent"

ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(4, 16).Value = max_volume
ws.Cells(4, 15).Value = max_ticker

Next ws

End Sub
