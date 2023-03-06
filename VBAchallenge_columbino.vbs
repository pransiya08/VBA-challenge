Sub stockchallenge()

Dim WS_Count As Integer
Dim I As Integer

'Set WS_Count equal to the number of worksheets in the active workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

'Begin the loop.
For I = 1 To WS_Count

    'activate current worksheet
    Worksheets(I).Activate

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"


    'identify the length of the rows used for this dataset
    Dim length As Long
    length = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (length)

    'setting up the variables to use for the looping
    Dim ticker, tickcount, j As Long
    ticker = 2
    tickcount = 2
    
    'setting up the variables to calculate for the yearly changes and stock volume
    Dim op, clo As Double 'setting up as Double due to decimal places
    Dim stock_vol As Double 'setting up as Double to have more bytes - Long causes Overflow error
    op = 0 'opening value
    clo = 0 'closing value
    stock_vol = 0 'stock volume

    'iteration to check for and print required values
    For j = 2 To length
  
        'to check if the current value is not the same as the one in the previous row
        If (Range("A" & j).Value <> Range("A" & (j - 1)).Value) Then
            'if it not the same, then it is a unique value of ticker
            Range("I" & ticker).Value = Range("A" & j).Value
        
            'getting the opening value of the ticker (in the row of the first instance)
            op = Range("C" & j).Value
        
            'counter for the next ticker
            ticker = ticker + 1
        
            'if its a new ticker, stock volume count will reset to 0
            stock_vol = 0
        End If
    
      
        'getting the closing value of the ticker in the row of the last instance
        clo = Range("F" & tickcount).Value
    
        'calculating and formatting the yearly changes
        Dim y_change As Double
        y_change = clo - op
        Range("J" & ticker - 1).Value = y_change
        'formatting it to have the currency format
        Range("J" & ticker - 1).NumberFormat = "$#,##0.00"
        'conditional formatting to apply if it's negative to be red, otherwise green
        If (Range("J" & ticker - 1).Value < 0) Then
            Range("J" & ticker - 1).Interior.ColorIndex = 3
        Else
            Range("J" & ticker - 1).Interior.ColorIndex = 4
        End If

        'as long as the ticker is the same as in the next row, it will just sum up the stock volume
        stock_vol = stock_vol + Range("G" & tickcount).Value
        Range("L" & ticker - 1).Value = stock_vol
        Range("L" & ticker - 1).NumberFormat = "0"
      
        'calculating and formatting the percent change to percentage format with up to 2 decimal places
        Range("K" & ticker - 1).Value = y_change / op
        Range("K" & ticker - 1).NumberFormat = "0.00%"
    
        'counter for the actual iteration to determine the last row of a unique ticker
        tickcount = tickcount + 1
    
    Next j
  
    'setting up the table headers of the Calculated Values
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest %Increase"
    Range("O3").Value = "Greatest %Decrease"
    Range("O4").Value = "Greatest Total Volume"
     
    'determining the length of the new table of unique tickers
    Dim lentab As Double
    lentab = Cells(Rows.Count, 9).End(xlUp).Row
        
    'setting up the For loop to determine the Calculated Values
    Dim k As Integer
    
    'finding the max increase% and printing it with the right formatting
    Range("Q2").Value = Application.WorksheetFunction.Max(Range("K:K").Value)
    Range("Q2").NumberFormat = "0.00%"
       
    'utilising the For loop to look for and print the matching value with the max %increase
    For k = 2 To lentab
        If (Range("K" & k).Value = Range("Q2").Value) Then
            Range("P2").Value = Range("I" & k).Value
        End If
    Next k
    
    'finding the max decrease% and printing it with the right formatting
    Range("Q3").Value = Application.WorksheetFunction.Min(Range("K:K").Value)
    Range("Q3").NumberFormat = "0.00%"
    
    'utilising the For loop to look for and print the matching value with the max %decrease
    For k = 2 To lentab
        If (Range("K" & k).Value = Range("Q3").Value) Then
            Range("P3").Value = Range("I" & k).Value
        End If
    Next k
    
    'finding the max stock volume
    Range("Q4").Value = Application.WorksheetFunction.Max(Range("L:L").Value)
    Range("Q4").NumberFormat = "0"
    'utilising the For loop to look for and print the matching value with the max stock volume
    For k = 2 To lentab
        If (Range("L" & k).Value = Range("Q4").Value) Then
            Range("P4").Value = Range("I" & k).Value
        End If
    Next k
    
    'auto-fitting the columns to make it more cleaner
    Worksheets(I).Columns.AutoFit

Next I

End Sub