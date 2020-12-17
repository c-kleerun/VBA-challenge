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
             
Next ws

End Sub


