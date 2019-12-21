Attribute VB_Name = "Module1"
Sub VbaStocks()

    'loop to go through all the sheets
        For Each ws In Worksheets
        
        'Colum headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Declare and intialize variables
        Dim Ticker_Name As String
        Dim Last_Row As Long
        
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        Dim Yearly_Open As Double
        Dim Yearly_Close As Double
        Dim Yearly_Change As Double
        
        Dim Previous_Amount As Long
        Previous_Amount = 2
        
        Dim Percent_Change As Double
        
        Dim Greatest_Increase As Double
        Greatest_Increase = 0
        
        Dim Greatest_Decrease As Double
        Greatest_Decrease = 0
        
        Dim Last_Row_Value As Long
        
        Dim Greatest_Total_Volume As Double
        Greatest_Total_Volume = 0
        
        
         ' Last row of the data
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'loop to last row
        For i = 2 To Last_Row

            ' Add to ticker total volume
            Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
            
            ' Check whether if we are in the same ticker name
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' if it is not set Ticker Name
                Ticker_Name = ws.Cells(i, 1).Value
                
                ' Print ticker name in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                ' Print ticker totL amoutn to the summary table
                ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                ' Reset ticker total
                Total_Ticker_Volume = 0


                ' Set yearly pen, yearly close and yearly change name
                Yearly_Open = ws.Range("C" & Previous_Amount)
                Yearly_Close = ws.Range("F" & i)
                Yearly_Change = Yearly_Close - Yearly_Open
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                ' find the percent change
                If Yearly_Open = 0 Then
                    Percent_Change = 0
                Else
                    Yearly_Open = ws.Range("C" & Previous_Amount)
                    Percent_Change = Yearly_Change / Yearly_Open
                End If
                ' Format column to add % sign and 2 decimal places
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change

                ' Conditional formatting highlight
                If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            
                ' Add one to the summary table Row
                Summary_Table_Row = Summary_Table_Row + 1
                Previous_Amount = i + 1
                End If
            Next i



            ' Greatest % increase, greatest % decrease and greatest total volume
            Last_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            ' Start loop for Ffinal results
            For i = 2 To Last_Row
            
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
        ' Format column to include % 2 decimal places
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
          
    
    Next ws
    
End Sub


