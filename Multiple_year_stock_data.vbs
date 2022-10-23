Option Explicit
Sub Stock_Checker_WorksheetLoop()

'This vba script will output required stock checker data points on each sheet by looping through all sheets

'Declare Current as a worksheet object variable
Dim a As Integer
Dim x As Integer

'Set variable for sheet number
a = Application.Worksheets.Count

'Loop through all of the worksheets in the active workbook
For x = 1 To a

Worksheets(x).Activate

'Begin script for each sheet

    'Set an initial variable for holding the ticker symbol
    Dim Ticker_Symbol As String
    
    'Set an initial variable for holding the total stock volume
    Dim Total_Volume As LongLong
    Total_Volume = 0
    
    'Counts the number of rows
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set variables for counters
    Dim i As Long
            
    'Set an initial variable for holding the first opening price and closing price at end of year for each Ticker Symbol
    Dim Ticker_Open As Double
    Dim Ticker_Close As Double
    Dim Ticker_Change As Double
    Dim Ticker_Pct_Change As Double
    
    'Set variables to hold greatest increase, decrease and total volume
    Dim Greatest_Increase As Double
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease As Double
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Volume As LongLong
    Dim Greatest_Volume_Ticker As String
    
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Volume = 0
               
    'Keep track of the location for each Ticker Symbol in the Summary Table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Loop through all rows of all ticker symbols
    For i = 2 To lastrow
    'For i = 2 To 71
            
        'Check if we are still within the same ticker symbol, if we are not...
        If Cells(i - 1, 1).Value <> Cells(i, 1) Then
            
            'Set the Ticker Symbol & Open
            Ticker_Symbol = Cells(i, 1).Value
            Ticker_Open = Cells(i, 3).Value
            
            'Add to Total Volume
            Total_Volume = Total_Volume + Cells(i, 7).Value
                        
                   
            'Print the next Ticker Symbol in the Summary Table
            Range("I" & Summary_Table_Row).Value = Ticker_Symbol
        
            'Print the next Ticker_Open in the Summary Table
            'Range("J" & Summary_Table_Row).Value = Ticker_Open
            
            Else
                Total_Volume = Total_Volume + Cells(i, 7).Value
                If Cells(i + 1, 1).Value <> Cells(i, 1) Then
                Ticker_Close = Cells(i, 6)
                
                'Print the next Ticker_Close in the Summary Table
                'Range("K" & Summary_Table_Row).Value = Ticker_Close
                
                Ticker_Change = Ticker_Close - Ticker_Open
                
                'Print Yearly Change in the Summary Table and color code
                Range("J" & Summary_Table_Row).Value = Ticker_Change
                
                    If Range("J" & Summary_Table_Row).Value > 0 Then
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 10
                       
                        ElseIf Range("J" & Summary_Table_Row).Value < 0 Then
                            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                            
                        Else
                            Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
                    
                    End If
                
                'Calculate Percent Change
                Ticker_Pct_Change = Ticker_Change / Ticker_Open
                
                'Print Yearly Percent Change in Summary Table
                Range("K" & Summary_Table_Row).Value = FormatPercent(Ticker_Pct_Change, 2)
                
                'Print Total_Volume
                Range("L" & Summary_Table_Row).Value = Total_Volume
                
                'Store Greatest Total Volume
                If Total_Volume > Greatest_Volume Then
                    Greatest_Volume = Total_Volume
                    Greatest_Volume_Ticker = Ticker_Symbol
                
                End If
            
                'Store Greatest Percent Increase - Percent Change
                If Ticker_Pct_Change > Greatest_Increase Then
                    Greatest_Increase = Ticker_Pct_Change
                    Greatest_Increase_Ticker = Ticker_Symbol
                    
                End If
                
                'Store Greatest Percent Decrease - Percent Change
                If Ticker_Pct_Change < Greatest_Decrease Then
                    Greatest_Decrease = Ticker_Pct_Change
                    Greatest_Decrease_Ticker = Ticker_Symbol
                    
                End If
            
                'Add one to the summary table now
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Set Total Volume back to 0
                Total_Volume = 0
        
            End If
            
        End If
        
    Next i
        'Write Bonus Cells
        Cells(2, 16).Value = Greatest_Increase_Ticker
        Cells(2, 17).Value = FormatPercent(Greatest_Increase, 2)
        Cells(3, 16).Value = Greatest_Decrease_Ticker
        Cells(3, 17).Value = FormatPercent(Greatest_Decrease, 2)
        Cells(4, 16).Value = Greatest_Volume_Ticker
        Cells(4, 17).Value = Greatest_Volume

Next x

End Sub







    


