' FOR CONDITIONAL FORMATTING - From Stack Overflow (https://stackoverflow.com/questions/27611260/what-are-the-rgb-codes-for-the-conditional-formatting-styles-in-excel)
'                    For 'Bad' red:
'                    Font Is: (156,0,6)
'                    Background Is: (255,199,206)
'
'                    For 'Good' green:
'                    Font Is: (0,97,0)
'                    Background Is: (198,239,206)
'
'                    For 'Neutral' yellow:
'                    Font Is: (156,101,0)
'                    Background Is: (255,235,156)
'
' I mostly used the credit card exercise we worked on in class, along with going over the lecture multiple times to get this figured out.
' However I also used Microsoft VBA documentation for info on how to format and other information (eg: FormatPercent, autofit, variable types - LongLong)
' I also used ExcelEasy (https://www.excel-easy.com/vba.html) for more information, examples and how to deal with VBA errors

Sub MultiYearStocks():
' LOOP THRU WORKSHEETS
    For Each ws In Worksheets
    
    'Variables
    Dim WorksheetName As String
    WorksheetName = ws.Name
    Dim TickerName As String
    Dim PriceStart As Double
    Dim PriceEnd As Double
    Dim Volume As LongLong
    Dim Summary_Table_Row As Integer
    Dim EndRowA As Long
    Dim CurrentRow As Long
    Dim YearlyChange As Double ' For Debugging only
    Dim PercentChange As Double ' For Debugging only
    Dim GreatIncrease As Double
    Dim GreatDecrease As Double
    Dim GreatTotalVolume As LongLong
    Dim GreatIncrease_Ticker As String
    Dim GreatDecrease_Ticker As String
    Dim GreatTotalVolume_Ticker As String
        
    'Set Initial values
    Volume = 0
    Summary_Table_Row = 2
    GreatIncrease = 0
    GreatDecrease = 0
    GreatTotalVolume = 0
        
    'Create summary columns
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
       
    'Find LastRow in column A
    EndRowA = ws.Range("A1").End(xlDown).Row
    
    'FOR DEBUGGING - Tell me what the last row is
    'MsgBox ("Column A last row = " & EndRowA)
        
        'Loop through all rows
        For CurrentRow = 2 To EndRowA
            
            'Check if ticker is the same as previous row
            If ws.Cells(CurrentRow, "A").Value <> ws.Cells(CurrentRow - 1, "A").Value Then
                
                'Set current ticker name
                TickerName = ws.Cells(CurrentRow, "A").Value
                
                'Set starting price
                PriceStart = ws.Cells(CurrentRow, "C").Value
                
                'Add first value to Volume
                Volume = Volume + ws.Cells(CurrentRow, "G").Value
                
                'Add ticker name to the summary table
                ws.Range("I" & Summary_Table_Row).Value = TickerName
                
                'FOR DEBUGGING - Tell me the Starting Price
                'MsgBox ("Starting Price = " & PriceStart)
                
                'Check if ticker changed in the next row
            ElseIf ws.Cells(CurrentRow, "A").Value <> ws.Cells(CurrentRow + 1, "A").Value Then
                
                'Set ending price
                PriceEnd = ws.Cells(CurrentRow, "F").Value
                
                'FOR DEBUGGING - Tell me the Ending Price
                'MsgBox ("Ending Price = " & PriceEnd)
                
                'Add new row volume to total stock volume
                Volume = Volume + ws.Cells(CurrentRow, "G").Value
                
                'FOR DEBUGGING - Tell me the Volume
                'MsgBox ("Volume = " & Volume)
                
                'Calculate Yearly Change and format
                ws.Range("J" & Summary_Table_Row).Value = (PriceEnd - PriceStart)
                    If ws.Range("J" & Summary_Table_Row).Value < 0 Then ' Conditional Formatting
                    ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 199, 206) 'Set cell background to red
                    Else
                    ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(198, 239, 206) 'Set cell background color to green
                    End If
                
               'FOR DEBUGGING - Calculation - Yearly Change
               'YearlyChange = Round(PriceEnd - PriceStart, 2)
               'MsgBox ("Ending Price = " & PriceEnd & vbCrLf & "Starting Price = " & PriceStart & vbCrLf & "Yearly Change = " & YearlyChange)
                
                'Percent Change Calculation with formatting
                ws.Range("K" & Summary_Table_Row).Value = FormatPercent(((PriceEnd - PriceStart) / PriceStart))
                
                'FOR DEBUGGING - Calculation - Percent Change
                'PercentChange = (((PriceEnd - PriceStart) / PriceStart)) * 100
                'PercentChange = Round((PercentChange), 2)
                'MsgBox ("Percent Change = " & PercentChange)
                
                'Add total volume
                ws.Range("L" & Summary_Table_Row).Value = Volume
                
                'Add new row to summary table
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset Volume counter
                Volume = 0
                
            Else
                'Add to ticker Volume
                Volume = Volume + ws.Cells(CurrentRow, "G").Value
                
            End If
        Next CurrentRow
        
        'Create 2nd loop to get overall greatest increase/decrease/volume
        For OverallSummary = 2 To EndRowA
            If ws.Cells(OverallSummary, "K").Value > GreatIncrease Then
                GreatIncrease_Ticker = ws.Cells(OverallSummary, "I").Value
                GreatIncrease = ws.Cells(OverallSummary, "K").Value
                
            ElseIf ws.Cells(OverallSummary, "K").Value < GreatDecrease Then
                GreatDecrease_Ticker = ws.Cells(OverallSummary, "I").Value
                GreatDecrease = ws.Cells(OverallSummary, "K").Value
           
            ElseIf ws.Cells(OverallSummary, "L").Value > GreatTotalVolume Then
                GreatTotalVolume_Ticker = ws.Cells(OverallSummary, "I").Value
                GreatTotalVolume = ws.Cells(OverallSummary, 12)

            End If
        Next OverallSummary
        
        'Put everything in overal summary
        ws.Cells(2, "P").Value = GreatIncrease_Ticker
        ws.Cells(3, "P").Value = GreatDecrease_Ticker
        ws.Cells(4, "P").Value = GreatTotalVolume_Ticker
        
        'Format all cells correctly
        ws.Cells(2, "Q").Value = FormatPercent(GreatIncrease)
        ws.Cells(3, "Q").Value = FormatPercent(GreatDecrease)
        ws.Cells(4, "Q").Value = Format(GreatTotalVolume, "Scientific")
        ws.Columns("J").NumberFormat = "0.00"
        ws.Columns("A:Z").AutoFit
        
    Next ws
    
End Sub
