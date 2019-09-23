Attribute VB_Name = "Module1"
Sub testonlargetable()

Dim ws As Worksheet

For Each ws In Worksheets

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim Ticker As String
    
    'Set volume
    Dim Stock_Volume As Double
    Stock_Volume = 0
    
    ' Keep track of the location for each ticker in the summary table
    Dim NewTicker_SummaryTable As Integer
    NewTicker_SummaryTable = 2
    
    ' Opening price for first ticker in the sheet
    Dim OpenPrice As Double
    OpenPrice = ws.Cells(2, 3).Value
  
    ' Adding labels
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly_Change"
    ws.Cells(1, 11).Value = "Percent_Change"
    ws.Cells(1, 12).Value = "Stock_Volume"
  
  
    ' Loops to get rest of the data
    For i = 2 To LastRow
    
        Dim ClosePrice As Double
        
        ' Check if we are still within the same stock
            
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ' Set the Values
        Ticker = ws.Cells(i, 1).Value
        ClosePrice = ws.Cells(i, 6).Value
    
            
        ' Calculations
        Yearly_Change = ClosePrice - OpenPrice
        Percent_Change = Yearly_Change / OpenPrice
        
       
        ' Add to the Stock Volume
          Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
          
        ' Print the Ticker Info in the Summary Table
    
        ws.Range("I" & NewTicker_SummaryTable).Value = Ticker
        ws.Range("J" & NewTicker_SummaryTable).Value = Yearly_Change
        ws.Range("K" & NewTicker_SummaryTable).Value = Percent_Change
        ws.Range("K" & NewTicker_SummaryTable).NumberFormat = "0.00%"
        ws.Range("L" & NewTicker_SummaryTable).Value = Stock_Volume
        
        ' Add one to the summary table row
        NewTicker_SummaryTable = NewTicker_SummaryTable + 1
        
        ' Recalculate open price
        If ws.Cells(i + 1, 3).Value <> 0 Then
        OpenPrice = ws.Cells(i + 1, 3).Value
        End If
        
        ' Reset the stock volume
          Stock_Volume = 0
    
        ' If the cell immediately following a row is the stock
        Else
    
          ' Add to the Volume total
          Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    FinalRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    'MsgBox (FinalRow)
    For i = 2 To FinalRow
        If ws.Cells(i, 10).Value > 0 Then
    
           ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10) < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
        
    ' Adding labels
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
   
    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K" & 2 & ":" & "K" & FinalRow))
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("I" & 2 & ":" & "I" & FinalRow), WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K" & 2 & ":" & "K" & FinalRow), 0))
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K" & 2 & ":" & "K" & FinalRow))
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P3").Value = WorksheetFunction.Index(ws.Range("I" & 2 & ":" & "I" & FinalRow), WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K" & 2 & ":" & "K" & FinalRow), 0))
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L" & 2 & ":" & "L" & FinalRow))
    ws.Range("P4").Value = WorksheetFunction.Index(ws.Range("I" & 2 & ":" & "I" & FinalRow), WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L" & 2 & ":" & "L" & FinalRow), 0))

Next ws

End Sub

