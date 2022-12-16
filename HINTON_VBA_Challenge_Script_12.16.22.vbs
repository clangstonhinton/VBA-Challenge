Sub VBA_Challenge()

' Set variable for Worksheet
Dim ws As Worksheet

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    ' Set initial variable Ticker Symbol
    Dim Ticker As String
    
    ' Set initial variable Total Volume
    Dim Volume_Total As Double
    
    ' Set initial variable for Yearly Change
    Dim Yearly_Change As Double
    
    ' Set initial variable for Start of Year Open
    Dim Start_of_Year_Open As Double
    
    ' Set initial variable for Percent Change
    Dim Percent_Change As Double
    
    ' Keep track of the location for each Ticker Symbol in the summary table
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    
    ' Determine the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Create headers for ticker, yearly change, percent change, and total stock volume
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    'Format Header text in bold font and wrap text
    ws.Range("I1:L1").Font.Bold = True
    ws.Range("I1:L1").WrapText = True
  
    'Format Percent Change column with 2 decimal places and Volume Total column numbers is comma separating 1000s
    ws.Columns("L").NumberFormat = "#,###"
    ws.Columns("K").NumberFormat = "0.00%"
    
    'Center Header text in columns J and K
    ws.Columns("J:K").HorizontalAlignment = xlCenter
  
    'Align Header text to right side in columns L and set column width to 15
    ws.Columns("L").HorizontalAlignment = xlRight
    ws.Columns("L").ColumnWidth = 15

    ' Assign initial variables
    Yearly_Change = 0
    Start_of_Year_Open = ws.Cells(2, 3).Value
    Percent_Change = 0
    Volume_Total = 0
        
    ' Keep track of the location of each stock ticker in ticker column
    Summary_Table_Row = 2

    ' Loop through year of stock data
    For i = 2 To LastRow
        
            'Check if stock ticker changes when moving to the next row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
                    ' Set stock ticker name
                    Ticker = ws.Cells(i, 1).Value
                
                    ' Add stock volume to total stock volume
                    Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
                    ' Subtract Start of Year Open from End of Year close to find the yearly change
                    Yearly_Change = ws.Cells(i, 6).Value - Start_of_Year_Open
     
                    ' Calculate Percent Change value
                    ' Avoid "divide by 0" error
                    If Start_of_Year_Open = 0 Then
                
                        Percent_Change = 0
                
                        Else
                               
                            ' Perform the percent change calculation
                            Percent_Change = Yearly_Change / Start_of_Year_Open
                
                    End If
                
                    ' Print stock ticker name in ticker column
                    ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                    ' Print yearly change in yearly change column
                    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                    ' Conditional formatting for yearly change
                    If Yearly_Change < 0 Then
                    
                            ' If negative change make interior color red
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                            ElseIf Yearly_Change > 0 Then
                    
                                    ' If positive change make interior color green
                                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

                
                    End If
                
                    ' Print percent change to percent change column
                    ws.Range("K" & Summary_Table_Row).Value = Format(Percent_Change, "Percent")
                
                    ' Print stock total volume amount to the stock volume column
                    ws.Range("L" & Summary_Table_Row).Value = Volume_Total
                
                    ' Add one  to summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                
                    ' Reset total stock volume
                    Volume_Total = 0
                
                    ' Reset start of year open price
                    Start_of_Year_Open = ws.Cells(i + 1, 3).Value
            
            Else
            
                    ' If cell immediately following row is the same stock ticker then just add stock volume to total stock volume
                    Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
            End If
                
        Next i

        ' Determine the last row in the Summary Table
        LastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row

        ' Create  Greatest Inc/Dec/Vol  table  headers
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Set variables for Greatest Inc/Dec/Vol table
        Dim Greatest_Percent_Increase_Ticker As String
        Dim Greatest_Percent_Increase As Double
        Dim Greatest_Percent_Decrease_Ticker As String
        Dim Greatest_Percent_Decrease As Double
        Dim Greatest_Total_Volume_Ticker As String
        Dim Greatest_Total_Volume As Double
        
        ' Define the variables for greatest percent increase, greatest percent decrease, and greatest total volume
        Greatest_Percent_Increase = WorksheetFunction.Max((Range(ws.Cells(2, 11), ws.Cells(LastRowSummary, 11))))
        Greatest_Percent_Decrease = WorksheetFunction.Min((Range(ws.Cells(2, 11), ws.Cells(LastRowSummary, 11))))
        Greatest_Total_Volume = WorksheetFunction.Max((Range(ws.Cells(2, 12), ws.Cells(LastRowSummary, 12))))
    
        ' Loop through summary table to match values with tickers
        For i = 2 To LastRowSummary
        
                    ' Find stock ticker with greatest percent increase
                    If ws.Cells(i, 11).Value = Greatest_Percent_Increase Then
            
                    ' Set stock ticker name
                    Greatest_Percent_Increase_Ticker = ws.Cells(i, 9).Value
                
                    ' Print stock ticker name in ticker column of Greatest Inc/Dec/Vol table
                    ws.Cells(2, 16).Value = Greatest_Percent_Increase_Ticker
                
                    ' Print Greatest Percent Increase amount in value column of Greatest Inc/Dec/Vol table
                    ws.Cells(2, 17).Value = Greatest_Percent_Increase
                
                Else
                    
            End If
       
            ' Find stock ticker with greatest percent decrease
            If ws.Cells(i, 11).Value = Greatest_Percent_Decrease Then
            
                    ' Set stock ticker name
                    Greatest_Percent_Decrease_Ticker = ws.Cells(i, 9).Value
                
                    ' Print stock ticker name in ticker column of Greatest Inc/Dec/Vol table
                    ws.Cells(3, 16).Value = Greatest_Percent_Decrease_Ticker
                
                    ' Print Greatest Percent Increase amount in value column of Greatest Inc/Dec/Vol table
                    ws.Cells(3, 17).Value = Greatest_Percent_Decrease
                
                Else
                    
            End If
            
            ' Find stock ticker with greatest total volume
            If ws.Cells(i, 12).Value = Greatest_Total_Volume Then
            
                    ' Set stock ticker name
                    Greatest_Total_Volume_Ticker = ws.Cells(i, 9).Value
                
                    ' Print stock ticker name in ticker column of Greatest Inc/Dec/Vol table
                    ws.Cells(4, 16).Value = Greatest_Total_Volume_Ticker
                
                    ' Print Greatest Total Volume amount in value column of Greatest Inc/Dec/Vol table
                    ws.Cells(4, 17).Value = Greatest_Total_Volume
                
                Else
                    
            End If

        Next i

        ' Reset ticker values
        Greatest_Percent_Increase_Ticker = ""
        Greatest_Percent_Decrease_Ticker = ""
        Greatest_Total_Volume_Ticker = ""

        'Format Greatest Inc/Dec/Vol table
        ws.Range("P1:Q1").Font.Bold = True
        ws.Columns("P").HorizontalAlignment = xlCenter
        ws.Columns("P").ColumnWidth = 8
        ws.Range("Q1").HorizontalAlignment = xlRight
        ws.Columns("Q").ColumnWidth = 20
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("O").ColumnWidth = 20
        ws.Cells(4, 17).NumberFormat = "#,###"

    Next ws

End Sub
