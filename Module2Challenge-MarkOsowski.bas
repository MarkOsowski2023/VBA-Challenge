Attribute VB_Name = "Module1"
Sub MultiYearStockData():

    ' ----------------------------------
    ' Looping through all worksheets
    ' ----------------------------------
    For Each ws In Worksheets
    
        ' ------------------------------
        ' Declaring and setting variables
        ' -------------------------------
        Dim lastrow As Long
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim Summary_Table_Row As Integer
        
        Summary_Table_Row = 2
        
        Dim Ticker_Name As String
        
        Dim Yearly_Change As Double
        
        Yearly_Change = 0
        
        Dim Percent_Change As Double
        
        Percent_Change = 0
        
        Dim Total_Stock_Volume As Double
        
        Total_Stock_Volume = 0
        
        Dim open_price As Double
        ' initial open price
        open_price = 24.44
        
        Dim close_price As Double
        
        close_price = 0
        
        Dim volume As Double
        
        ' ---------------------------------
        ' Adding column headers
        ' ---------------------------------
        ws.Cells(1, 9).Value = "Ticker"
        
        ws.Cells(1, 10).Value = "Yearly Change"
        
        ws.Cells(1, 11).Value = "Percent Change"
        
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' ------------------------------------
        ' Loop through row 2 to lastrow
        ' ------------------------------------
        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            
                Ticker_Name = ws.Cells(i, 1).Value
                
                ws.Cells(Summary_Table_Row, 9).Value = Ticker_Name
                
                close_price = ws.Cells(i, 6).Value
                
                
                
                volume = ws.Cells(i, 7).Value
                
                
                Yearly_Change = Yearly_Change + (close_price - open_price)
                
                Percent_Change = Percent_Change + ((close_price - open_price) / open_price)
                
                Total_Stock_Volume = Total_Stock_Volume + volume
                
                ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
                
                ws.Cells(Summary_Table_Row, 11).Value = FormatPercent(Percent_Change)
                
                ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Yearly_Change = 0
                
                Percent_Change = 0
                
                Total_Stock_Volume = 0
                
                open_price = ws.Cells(i + 1, 3).Value
                
                close_price = 0
                
                
            Else
            
                
                volume = ws.Cells(i, 7).Value
        
                        
                Total_Stock_Volume = Total_Stock_Volume + volume
                
                
                
                
                
                
            End If
            
            ' -------------------------------------
            ' Formatting Yearly Change Column
            ' -------------------------------------
            If ws.Cells(Summary_Table_Row, 10).Value >= 0 Then
                
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                
            Else
                
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                
            End If
            
        Next i
        
            
    
    Next ws
    
End Sub
