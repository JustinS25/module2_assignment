Attribute VB_Name = "Module1"
Sub stocks():
    
    'need string variable to output ticker value and need i counter for loop to be long variable as integers can't store enough values for the entire worksheet
    Dim ticker As String
    Dim i As Long
    Dim counter As Integer
    
    'row counter variable stored as long because of too much stock info for integer to hold
    Dim row_count As Long
    
    'Loops through each worksheet in the workbook
    For Each ws In ActiveWorkbook.Worksheets
    
    'Below statements set the names of the cells and their width
    ws.Range("I1").Value = "Ticker"
    ws.Range("I1").ColumnWidth = 10
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("J1").ColumnWidth = 15
    ws.Range("K1").Value = "Percent Change"
    ws.Range("K1").ColumnWidth = 15
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("L1").ColumnWidth = 20
    
    'counts number of rows of stock data. Need this to determine end of for loop
    row_count = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Below loop runs through entire data set
    For i = 2 To row_count + 1
    
        'Below if statement determines closing value, yearly change, percent change, and total stock volume
        If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) And ws.Cells(i - 1, 1) <> ws.Cells(1, 1) Then
            closer = ws.Cells(i - 1, 6)
            ws.Cells(counter + 1, 10) = closer - opener
            ws.Cells(counter + 1, 11) = ((closer - opener) / opener) * 100 & "%"
            ws.Cells(counter + 1, 12) = volume
            'Resetting total stock volume so next stock doesn't add on to previous
            volume = 0
            'Below if statement changes color of positive and negative percentages to green and red, respectively
            If ws.Cells(counter + 1, 10) >= 0 Then
                ws.Cells(counter + 1, 10).Interior.Color = RGB(0, 255, 0)
            Else
                ws.Cells(counter + 1, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
        End If
        
        'Below if statement determines ticker string value, opening value, and volume of stocks
        If ticker <> ws.Cells(i, 1) Then
            counter = counter + 1
            ticker = ws.Cells(i, 1)
            ws.Cells(counter + 1, 9) = ticker
            opener = ws.Cells(i, 3)
            volume = ws.Cells(i, 7) + volume
        'Else determines that if the ticker is the same as the previous cell, we simply add the volume of the same tickers to get the overall sum
        Else
            volume = ws.Cells(i, 7) + volume
        End If

    Next i
    
    'Below resets the counter to 0 so we can reset the counter when we go to the next page
    counter = 0
    
    'Below statements set the names of the cells and their width
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("O2").ColumnWidth = 20
    ws.Range("P1").Value = "Ticker"
    ws.Range("P1").ColumnWidth = 15
    ws.Range("Q1").Value = "Value"
    ws.Range("Q1").ColumnWidth = 15
    
    'counts number of rows of stock data. Need this to determine end of for loop
    Dim row_count2 As Integer
    row_count2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'New loop to go through percentages to determine greatest % increase and decrease
    'Below loop runs through entire data set
    For i = 2 To row_count2
        
            'if statement to see if increase is bigger than previous, replace the old biggest with the new biggest
            If ws.Cells(i, 11) > big_inc Then
                big_inc = ws.Cells(i, 11)
                ticker_inc = ws.Cells(i, 9)
            End If
            'if statement to see if decrease is bigger than previous, replace the old biggest with the new biggest
            If ws.Cells(i, 11) < big_dec Then
                big_dec = ws.Cells(i, 11)
                ticker_dec = ws.Cells(i, 9)
            End If
            'if statement to see if the total volume is greater than the previous, replace the old biggest with the new biggest
            If ws.Cells(i, 12) > big_vol Then
                big_vol = ws.Cells(i, 12)
                ticker_vol = ws.Cells(i, 9)
            End If
        'Below statement allows loop to go to cell below
        Next i
        
    'Below statements set the names of the cells
    ws.Range("P2").Value = ticker_inc
    ws.Range("P3").Value = ticker_dec
    ws.Range("P4").Value = ticker_vol
    ws.Range("Q2").Value = big_inc * 100 & "%"
    ws.Range("Q3").Value = big_dec * 100 & "%"
    ws.Range("Q4").Value = big_vol
    
    'Below resets the values to 0 before going to next worksheet page. Otherwise we could have stocks from previous pages leak into otther pages
    ticker_inc = 0
    ticker_dec = 0
    ticker_vol = 0
    big_inc = 0
    big_dec = 0
    big_vol = 0
    
    Next ws

End Sub

