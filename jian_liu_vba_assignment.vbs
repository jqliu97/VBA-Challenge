Attribute VB_Name = "Module1"
Sub alphabetical_testing()
Dim Ticker_Symbol As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double

For Each ws In Worksheets

Summary_Table = 2
LastRow = ws.Cells(ws.Cells.Rows.Count, "A").End(xlUp).Row
startrow = 2
Start = 2
For i = startrow To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_Symbol = ws.Cells(i, 1).Value
        Yearly_Change = (ws.Cells(i, 6) - ws.Cells(Start, 3))
        Percent_Change = (Yearly_Change / ws.Cells(Start, 3))
        Start = i + 1
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        ws.Cells(Summary_Table, 11).Value = Ticker_Symbol
        ws.Cells(Summary_Table, 12).Value = Yearly_Change
        ws.Cells(Summary_Table, 13).Value = Percent_Change
        ws.Cells(Summary_Table, 13).NumberFormat = "0.00%"
        ws.Cells(Summary_Table, 14).Value = Total_Stock_Volume
        
        Select Case Yearly_Change
                    Case Is > 0
                        ws.Range("L" & Summary_Table).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("L" & Summary_Table).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("L" & Summary_Table).Interior.ColorIndex = 0
                End Select
        
        Summary_Table = Summary_Table + 1
    
Yearly_Change = 0
Percent_Change = 0
Total_Stock_Volume = 0
        
    Else

        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    End If
Next i

Dim Max_Ticker As String
Dim Min_Ticker As String
Dim Max_Total_Ticker As String
lastrow_M = ws.Cells(ws.Cells.Rows.Count, "M").End(xlUp).Row
MaxValue = 0
MinValue = 0
MaxTotal = 0
For j = 2 To lastrow_M
    If ws.Cells(j, 13).Value > MaxValue Then
        MaxValue = ws.Cells(j, 13).Value
        Max_Ticker = ws.Cells(j, 11).Value
        MaxRow = j
    End If
    If ws.Cells(j, 13).Value < MinValue Then
        MinValue = ws.Cells(j, 13).Value
        Min_Ticker = ws.Cells(j, 11).Value
        MinRow = j
    End If
    If ws.Cells(j, 14).Value > MaxTotal Then
        MaxTotal = ws.Cells(j, 14).Value
        Max_Total_Ticker = ws.Cells(j, 11).Value
        MaxTotalRow = j
    End If
Next j

    ws.Range("Q2").Value = Max_Ticker
    ws.Range("R2").Value = MaxValue
    ws.Range("R2").NumberFormat = "0.00%"
    ws.Range("Q3").Value = Min_Ticker
    ws.Range("R3").Value = MinValue
    ws.Range("R3").NumberFormat = "0.00%"
    ws.Range("Q4").Value = Max_Total_Ticker
    ws.Range("R4").Value = MaxTotal
    
    ws.Range("K1").Value = "Ticker"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "Percent Change"
    ws.Range("N1").Value = "Total Stock Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"

Next ws



End Sub
