# stock-analysis

## Overview

##### Background

In Module 2, we created a workbook to help Steve analyze stock data for his parents. He is now able to run analysis with the push of a button to show a summary of the tickers by volume and total return percent. 

##### Purpose

The purpose of this analysis is to refactor the code previously used to analyze the stock data to run the same information in a more efficient way. I will be comparing how long the original code takes to execute in comparison to the new refactored code. We want to see if there is a way to cut down the run time as Steve gathers more data.  

## Results

###### Original Code

The original code for module 2 is below. The pictures below are screenshots of the message box from the original code. The macro for 2017 ran in .2988 seconds and the one for 2018 ran in .2949 seconds. The original code ran reltively fast but as Steve collects more data and has more to analyze, he will need a code that can create the same outputs in a faster amount of time.

<img width="224" alt="VBA_Challenge_2017 - original code" src="https://user-images.githubusercontent.com/107590196/176442400-c576a050-9739-48b6-a6ac-ab0ca7b5b078.png">

<img width="251" alt="VBA_Challenge_2018 - original code" src="https://user-images.githubusercontent.com/107590196/176442415-47b79d48-46b6-4a69-8ed2-0946516727bc.png">





    Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("Whatyear would you like to run the anlaysis on?")
    
    startTime = Timer

     '1) Format the output sheet on All Stocks Analysis worksheet
       Worksheets("All Stocks Analysis").Activate
       Range("A1").Value = "All Stocks (" + yearValue + ")"
       'Create a header row
       Cells(3, 1).Value = "Ticker"
       Cells(3, 2).Value = "Total Daily Volume"
       Cells(3, 3).Value = "Return"

       '2) Initialize array of all tickers
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
       
       '3a) Initialize variables for starting price and ending price
       Dim startingPrice As Single
       Dim endingPrice As Single
       
       '3b) Activate data worksheet
       Worksheets(yearValue).Activate
       
       '3c) Get the number of rows to loop over
       RowCount = Cells(Rows.Count, "A").End(xlUp).Row

       '4) Loop through tickers
       For i = 0 To 11
           ticker = tickers(i)
           totalVolume = 0
           '5) loop through rows in the data
           Worksheets(yearValue).Activate
           For j = 2 To RowCount
               '5a) Get total volume for current ticker
               If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

     Next i


        'Formatting
        Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit

        dataRowStart = 4

        dataRowEnd = 15

        For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green

            Cells(i, 3).Interior.Color = vbGreen


        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red

            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color

            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i

      endTime = Timer
            MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

###### Refactored Code

I was able to create a refactored code and run it successfully. The new code cut down the run time significantly by two-thirds down to .09375 for both 2017 an 2018.

<img width="224" alt="VBA_Challenge_2017 - original code" src="https://user-images.githubusercontent.com/107590196/176444099-6a36c988-4f2f-4262-bf20-8fb90d12dfa0.png">

<img width="251" alt="VBA_Challenge_2018 - original code" src="https://user-images.githubusercontent.com/107590196/176444127-d7c08eb6-5912-41db-bbb6-c2fe60872b20.png">

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

    '2a) Create a for loop to initialize the tickerVolumes to zero.
 
        For i = 0 To 11
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
        Next i

    '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount

    '3a) Increase volume for current ticker
      'increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  it is the first row then assign correct starting price
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
        'If it is the last row then assign correc ending price
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         'If the next ticker isn't the same as the last then increase the tickerIndex
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

    Next i

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

## Summary

##### Advantages

##### Disadvantages


