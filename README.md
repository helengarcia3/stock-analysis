# stock-analysis

## Overview

##### Background
In Module 2, we created a workbook to help Steve analyze stock data for his parents. He is now able to run analysis with the push of a button to show a summary of the tickers by volume and total return. 

##### Purpose
The purpose of this analysis is to refactor the code previously used to analyze the data and to run the same information in a more efficient way. 

## Results
The pictures below are screenshots from  below shows the reults for the first

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

<img width="224" alt="VBA_Challenge_2017 - original code" src="https://user-images.githubusercontent.com/107590196/176432839-455243e2-c67b-41ca-b088-cd40d867063e.png">

<img width="251" alt="VBA_Challenge_2018 - original code" src="https://user-images.githubusercontent.com/107590196/176432846-a71c1765-b30a-4b2a-8b65-f49d3412158a.png">

ro for running year 2017. The Macro ran in .2988 seconds.

## Summary

##### Advantages

##### Disadvantages


