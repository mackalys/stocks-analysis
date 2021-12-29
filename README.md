# Stock Analysis with VBA

## Overview of Project

### Purpose
#### The purpose of this project was to help Steve analyize the stock market for a renewable energy company his parents can invest in.  After analyizing the data, we made the workbook more interactive so Steve could have a better visual representation of the stock market that he could show his parents. 

## Results
### Overall
#### Based on the results from 2017 and 2018, Steve should suggest to his parents to invest in stocks from either RUN or ENPH, favoring RUN just because their performance is still just slightly better than ENPH. Looking at all the stocks, the market's performance in 2017 did much better than in 2018 although the reason is unknown.
### Comparing Code
#### When comparing the original script's and the refactored script's performances, the refactored script executed the commands quicker than the orginal. If you look at the images below, you can see they both ran under a second. The original script ran just slightly slower than the refactored script.
![VBA_Challenge_2017](https://github.com/mackalys/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/mackalys/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png)
#### If you look at the refactored code (shown below), you can see that it was just condensed from the original code, leading it to run quicker than before. In the refactored code, both the loops and the formatting is condensed into less loops and added the formatting into the main part of the code ratgher than in a seperate macro.

##### Original
    Sub AllStocksAnalysis()
        Dim startTime As Single
        Dim endTime As Single

          yearValue = InputBox("What year would you like to run the analysis on?")

          startTime = Timer

      '1) Format the output sheet on the "All Stocks Analysis" worksheet.
          Worksheets("All Stocks Analysis").Activate
              Range("A1").Value = "All Stocks (" + yearValue + ")"
        
              'Create a header row
              Cells(3, 1).Value = "Ticker"
              Cells(3, 2).Value = "Total Daily Volume"
              Cells(3, 3).Value = "Return"

      '2) Initialize an array of all tickers.
          Dim tickers(11) As String
    
          tickers(0) = "AY"
          tickers(1) = "CSIQ"
          tickers(2) = "DQ"
          tickers(3) = "ENPH"
          tickers(4) = "FSLR"
          tickers(5) = "HASI"
          tickers(6) = "JKS"
          tickers(7) = "RUN"
          tickers(8) = "SEDG"
          tickers(9) = "SPWR"
          tickers(10) = "TERP"
          tickers(11) = "VSLR"
    
      '3a) Prepare for the analysis of tickers.
        Dim startingPrice As Single
        Dim endingPrice As Single

      '3b) Initialize variables for the starting price and ending price.
          Sheets(yearValue).Activate

      '3c) Activate the data worksheet.
          RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
      '4) Find the number of rows to loop over.
          For i = 0 To 11
              ticker = tickers(i)
              totalVolume = 0

              '5) Loop through the tickers.
              Sheets(yearValue).Activate
              For j = 2 To RowCount

                  '5a) Loop through rows in the data.
                  If Cells(j, 1).Value = ticker Then
                      totalVolume = totalVolume + Cells(j, 8).Value
                  End If
    
                  '5b) Find the total volume for the current ticker.
                  If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                      startingPrice = Cells(j, 6).Value
                  End If
    
                  '5c) Find the starting price for the current ticker.
                  If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                      endingPrice = Cells(j, 6).Value
                  End If
        
              Next j
    
          '6) Output the data for the current ticker.
              Worksheets("All Stocks Analysis").Activate
              Cells(4 + i, 1).Value = ticker
              Cells(4 + i, 2).Value = totalVolume
              Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
              
          Next i

              endTime = Timer
              MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

      End Sub

      Sub formatAllStockAnalysisTable()

          'Formatting
          Worksheets("All Stocks Analysis").Activate
          Range("A3:C3").Font.Bold = True
          Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
          Range("A3:C3").Font.Color = vbBlue
          Range("A3:C3").Borders.Color = vbBlue
          Range("A3:C3").Font.ColorIndex = 22
          Range("A3:C3").Borders.ColorIndex = 12
          Range("B4:B15").NumberFormat = "$#,##0"
          Range("C4:C15").NumberFormat = "0.00%"
          Columns("B").AutoFit
    
        'Conditional Formatting
          If Cells(4, 3) > 0 Then
              'Color the cell Green
              Cells(4, 3).Interior.Color = vbGreen
        
          ElseIf Cells(4, 3) < 0 Then
              'Color the cell red
              Cells(4, 3).Interior.Color = vbRed
        
          Else
              'Clear the cell color
              Cells(4, 3).Interior.Color = xlNone
    
          End If

          dataRowStart = 4
          dataRowEnd = 15
          For i = dataRowStart To dataRowEnd
    
              If Cells(i, 3) > 0 Then
                
                  'Change cell color to Green
                  Cells(i, 3).Interior.Color = vbGreen
            
              ElseIf Cells(i, 3) < 0 Then
        
                  'Change cell color to Red
                  Cells(i, 3).Interior.Color = vbRed
            
              Else
            
                  'Clear the cell color
                  Cells(i, 3).Interior.Color = xlNone
            
              End If
        
          Next i

      End Sub

##### Refactored
      'Code Source https://github.com/caseychen3605/stock-analysis
      Sub AllStocksAnalysisRefactored()
          Dim startTime As Single
          Dim endTime  As Single

          yearValue = InputBox("What year would you like to run the analysis on?")

          startTime = Timer
    
          'Format the output sheet on All Stocks Analysis worksheet
          Worksheets("All Stocks Analysis").Activate
    
          Range("A1").Value = "All Stocks (" + yearValue + ")"
    
          'Create a header row
          Cells(3, 1).Value = "Ticker"
          Cells(3, 2).Value = "Total Daily Volume"
          Cells(3, 3).Value = "Return"

          'Initialize array of all tickers
          Dim tickers(12) As String
    
          tickers(0) = "AY"
          tickers(1) = "CSIQ"
          tickers(2) = "DQ"
          tickers(3) = "ENPH"
          tickers(4) = "FSLR"
          tickers(5) = "HASI"
          tickers(6) = "JKS"
          tickers(7) = "RUN"
          tickers(8) = "SEDG"
          tickers(9) = "SPWR"
          tickers(10) = "TERP"
          tickers(11) = "VSLR"
    
          'Activate data worksheet
          Worksheets(yearValue).Activate
    
          'Get the number of rows to loop over
          RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
          '1a) Create a ticker Index
          tickerIndex = 0

          '1b) Create three output arrays
          Dim tickerVolumes(12) As Long
          Dim tickerStartingPrices(12) As Single
          Dim tickerEndingPrices(12) As Single
    
          ''2a) Create a for loop to initialize the tickerVolumes to zero.
          For m = 0 To 11
              tickerVolumes(m) = 0
              tickerStartingPrices(m) = 0
              tickerEndingPrices(m) = 0
          Next m
        
          ''2b) Loop over all the rows in the spreadsheet.
          For m = 2 To RowCount
    
              '3a) Increase volume for current ticker
              tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(m, 8).Value
        
              '3b) Check if the current row is the first row with the selected tickerIndex.
              'If  Then
              If Cells(m, 1).Value = tickers(tickerIndex) And Cells(m - 1, 1).Value <> tickers(tickerIndex) Then
                  tickerStartingPrices(tickerIndex) = Cells(m, 6).Value
              'End If
              End If
        
              '3c) check if the current row is the last row with the selected ticker
              'If the next row’s ticker doesn’t match, increase the tickerIndex.
              'If  Then
              If Cells(m, 1).Value = tickers(tickerIndex) And Cells(m + 1, 1).Value <> tickers(tickerIndex) Then
                  tickerEndingPrices(tickerIndex) = Cells(m, 6).Value
              'End If
              End If

                  '3d Increase the tickerIndex.
                  'If Then
                  If Cells(m, 1).Value = tickers(tickerIndex) And Cells(m + 1, 1).Value <> tickers(tickerIndex) Then
                      tickerIndex = tickerIndex + 1
                  'End If
                  End If
    
          Next m
    
          '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
          For m = 0 To 11
        
              Worksheets("All Stocks Analysis").Activate
              'Output
              Cells(4 + m, 1).Value = tickers(m)
              Cells(4 + m, 2).Value = tickerVolumes(m)
              Cells(4 + m, 3).Value = tickerEndingPrices(m) / tickerStartingPrices(m) - 1
        
        
          Next m
    
          'Formatting
          Worksheets("All Stocks Analysis").Activate
          Range("A3:C3").Font.FontStyle = "Bold"
          Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
          Range("B4:B15").NumberFormat = "#,##0"
          Range("C4:C15").NumberFormat = "0.0%"
          Columns("B").AutoFit

          dataRowStart = 4
          dataRowEnd = 15

          For m = dataRowStart To dataRowEnd
        
              If Cells(m, 3) > 0 Then
            
                  Cells(m, 3).Interior.Color = vbGreen
            
              Else
        
                  Cells(m, 3).Interior.Color = vbRed
            
              End If
        
          Next m
 
          endTime = Timer
          MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

      End Sub

## Summary
#### Some of the advantages to refactoring code is just like what happened in this example: the code ran faster. Refactoring the code means that we have the oppurtunity to shape up the code where we need to and condense it. A disadvantage to refactoring is that you are kind of fixing something that isn't broken. The only thing witht that is if something works why change it? Going back to the advange though, it would be better to have code run more effectively if possible. In the case of this example, refactoring the code made the analysis run faster for Steve to work with. The only real con to refactoring was taking more time to make the code run faster rather than just leaving the working code the way it was.
