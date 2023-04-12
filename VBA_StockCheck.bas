Attribute VB_Name = "Module1"
Sub StockCheck():

'*****how to loop through workbook/sheets?
For Each ws In Worksheets

'set up variables
Dim Ticker As String
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Price_Change As Double
Dim Percent_Change As Double
Dim Total_Stock As Double
Dim Column_Count As Double
Dim Table_Row As Integer

'second loop variables
Dim Max_Change As Double
Dim Min_Change As Double
Dim Max_Total As Double
Dim Max_Ticker1 As String
Dim Max_Ticker2 As String
Dim Max_Ticker3 As String

'set variable values
Column_Count = WorksheetFunction.CountA(Columns(1))
Table_Row = 2
Total_Stock = 0
Opening_Price = ws.Cells(2, 3).Value
Yearly_Change = 0
Percent_Change = 0

'second loop variables
Max_Change = 0
Min_Change = 0
Max_Total = 0

'label columns for collected data
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'set up loop for data length
For X = 2 To Column_Count
    
    'set loop to gather data
    If ws.Cells(X + 1, 1).Value <> ws.Cells(X, 1).Value Then
        
        Ticker = ws.Cells(X, 1)
        Closing_Price = ws.Cells(X, 6)
        Total_Stock = Total_Stock + ws.Cells(X, 7).Value
        Yearly_Change = Closing_Price - Opening_Price
        Percent_Change = (Yearly_Change / Opening_Price) * 100
        
            'format positive and negative changes
            If ws.Range("J" & Table_Row).Value < 0 Then
                
                ws.Range("J" & Table_Row).Interior.ColorIndex = 3
                'ws.Range("K" & Table_Row).Interior.ColorIndex = 3
            Else
            
                ws.Range("J" & Table_Row).Interior.ColorIndex = 4
                'ws.Range("K" & Table_Row).Interior.ColorIndex = 4
            End If
            
        ws.Range("I" & Table_Row).Value = Ticker
        ws.Range("L" & Table_Row).Value = Total_Stock
        ws.Range("J" & Table_Row).Value = Yearly_Change
        ws.Range("K" & Table_Row).Value = FormatNumber(Percent_Change) + "%"
        Table_Row = Table_Row + 1
            
        Opening_Price = ws.Cells(X + 1, 3).Value
        Total_Stock = 0
    
    Else
        Total_Stock = Total_Stock + ws.Cells(X, 7).Value
        
    End If
    
Next X

    'loops data collected
    For Y = 2 To WorksheetFunction.CountA(ws.Columns(9))
        
        If Max_Change > ws.Cells(Y, 11).Value Then
            
            ws.Cells(2, 17).Value = Max_Change
            ws.Cells(2, 16).Value = Max_Ticker1
        
        Else
        
            Max_Change = ws.Cells(Y, 11).Value
            Max_Ticker1 = ws.Cells(Y, 9).Value
            
        End If
    
        If Min_Change < ws.Cells(Y, 11).Value Then
            
            ws.Cells(3, 17).Value = Min_Change
            ws.Cells(3, 16).Value = Max_Ticker2
        
        Else
        
            Min_Change = ws.Cells(Y, 11).Value
            Max_Ticker2 = ws.Cells(Y, 9).Value
            
        End If
        
        If Max_Total > Cells(Y, 12).Value Then
            
            ws.Cells(4, 17).Value = Max_Total
            ws.Cells(4, 16).Value = Max_Ticker3
        
        Else
        
            Max_Total = ws.Cells(Y, 12).Value
            Max_Ticker3 = ws.Cells(Y, 9).Value
            
        End If
        
    Next Y

Next ws

End Sub
