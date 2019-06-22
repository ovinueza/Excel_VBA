Attribute VB_Name = "Module1"
Sub Ticker_3()

'Loops/Activates the code for all worksheets
    For Each WS In Worksheets
    WS.Activate
    
'Define Ticker
    Dim Ticker_Name As String
'Define Ticker Total
    Dim Ticker_Total As Double
    Ticker_Total = 0
'Table of contents starts at row 2
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
'Last Row of table
    Dim lRow As Double
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
'Defines the values for opening, close prices and % changes
    Dim year_open_price As Double
    Dim year_close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    year_open_price = Cells(2, 3)
'Defines the Values for Column Q
    Dim GreatInc As Double
    Dim GreatDec As Double
    Dim GreatVol As Double
       
    
    
'Column Titles Lables
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
'Row O Title Lables
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    

'For Loop i starts at 2 until the last row with data
    For i = 2 To lRow

'If ticker symbol on current row is the same as the next row
    '
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
    
    Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
        If (year_open_price = 0) Then
        year_open_price = Cells(i + 1, 3).Value
        End If
    
    Else
        'If ticker symbol on the next row is different, capture ticker name, ticker total, yearly change values for the current ticker
        Ticker_Name = Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = Ticker_Name
        
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
        Range("L" & Summary_Table_Row).Value = Ticker_Total
    
        year_close_price = Cells(i, 6).Value
          
        yearly_change = year_close_price - year_open_price
        Range("J" & Summary_Table_Row).Value = yearly_change
              
    
    If (year_open_price = 0) Then
    Cells(Summary_Table_Row, 11).Value = 0
    
    'Only apply percent change calculation if year open price is not 0
    Else
    percent_change = (year_close_price - year_open_price) / year_open_price

    Range("K" & Summary_Table_Row).Value = FormatPercent(percent_change)

    End If
    
            'Color formatting for yearly change column
            If (Cells(Summary_Table_Row, 10).Value > 0 Or Cells(Summary_Table_Row, 10).Value = 0) Then
                Cells(Summary_Table_Row, 10).Interior.ColorIndex = 10
            ElseIf Cells(Summary_Table_Row, 10).Value < 0 Then
                Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                
            End If
    

    Cells(Summary_Table_Row, 12).Value = Ticker_Total
    
    'Indicate next row for next i to start with
    Summary_Table_Row = Summary_Table_Row + 1

    'Reset ticker total value to 0 for the next i
    
    Ticker_Total = 0
    
    year_open_price = Cells(i + 1, 3)
        
    End If


    Next i
    
    'MAX and MIN VALUES
  'Look through each rows to find the greatest value and its associate ticker

'Looks for the Max Value in Column K
    GreatInc = WorksheetFunction.Max(Range("K:K"))
    Range("Q2").Value = GreatInc
'Looks for the Min Value in Column K
    GreatDec = WorksheetFunction.Min(Range("K:K"))
    Range("Q3").Value = GreatDec
'Looks for the Max Value in Column L
    GreatVol = WorksheetFunction.Max(Range("L:L"))
    Range("Q4").Value = GreatVol
'Looks up the value on Q and retrives the Ticker
    Range("P2") = "=Index(I:I,match(Q2,K:K, 0))"
    Range("P3") = "=Index(I:I,match(Q3,K:K, 0))"
    Range("P4") = "=Index(I:I,match(Q4,L:L, 0))"
'Formats to percentage Values on Q2 & Q3
    Range("Q2:Q3").NumberFormat = "0.00%"
    
'Autofit the column for all ranges
        Range("I:Q").EntireColumn.AutoFit

        
    Next WS
    
End Sub
