Sub stock_analysis()
   
'to go through all the sheets
For Each ws In Worksheets
        
        
'Setting Variables
Dim Ticker As String
Dim Yearly_change As Double
Dim Percent_change As Double
Dim Total_stock_volume As Double
Dim Open_price As Double
Dim Close_price As Double
Dim LR As Double
Dim i As Long
Dim pointer As Long


Dim Summary_table As Double

'inserting header for the ST
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Change"
Columns("I:L").AutoFit

'Setting initial values for the var
Summary_table = 2
Total_stock_volume = 0
Open_price = 0
Close_price = 0
Ticker = ""

LR = ws.UsedRange.Rows.Count
        
            For i = 2 To LR
            
            Ticker = ws.Cells(i, 1).Value
            'if current price is zero then
            If Open_price = 0 Then
            Open_price = ws.Cells(i + 1, 3).Value
            ws.Cells(Summary_table, "I").Value = ws.Cells(i, 1).Value
            
            End If
            
            If ws.Cells(i, 1) <> ws.Cells(i - 1, 1).Value Then
            Open_price = ws.Cells(i, 3).Value
                
            End If
                  
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
         
            Else
            Close_price = ws.Cells(i, 6).Value
                
            Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
            Yearly_change = Close_price - Open_price
            'Percent_change = Round((Yearly_change / Open_price * 100), 2)
            Percent_change = (Yearly_change / Open_price * 100)
            
            
            If Total_stock_volume = 0 Then
            ws.Cells(Summary_table, "I").Value = ws.Cells(i, 1).Value
            ws.Cells(Summary_table, "J").Value = 0
            ws.Cells(Summary_table, "K").Value = "0%"
            ws.Cells(Summary_table, "L").Value = 0
                
            Else
            ws.Cells(Summary_table, "I").Value = ws.Cells(i, 1).Value
            ws.Cells(Summary_table, "J").Value = Yearly_change
            ws.Cells(Summary_table, "K").Value = "%" & Percent_change
            ws.Cells(Summary_table, "L").Value = Total_stock_volume
            
            End If
                
            If Yearly_change > 0 Then
            ws.Cells(Summary_table, "J").Interior.ColorIndex = 4
            ElseIf Yearly_change < 0 Then
            ws.Cells(Summary_table, "J").Interior.ColorIndex = 3
            
             End If
                
            'Reset variables for new stock ticker
            Total_stock_volume = 0
            Open_price = 0
            Close_price = 0
            Yearly_change = 0
            Summary_table = Summary_table + 1
            
            End If
            
            
    Next i
        
       
    Next ws
    
End Sub
