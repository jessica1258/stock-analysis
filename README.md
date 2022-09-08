# Stock Analysis Results
## Overview of Project

The purpose of this analysis is to analyze stocks for Steve's parents. It is a priority that the code and approach be able to work efficiently in order expand the analysis to additional stocks in the market in the future.

## Results

For the twelve stocks evaluated, all had trading volume of 80 million shares or more during 2018 which is sufficient transactions to ensure outlier prices do not result in incorrect conclusions. Of the twelve stocks, two had positive results. These are ENPH and RUN, which returned 81.9% and 84.0%, respectively.

These results were obtained using VBA code that was design to add volume for each individual stock over for all data for a year selected, 2018 in this case. In addition the annual return was calculated by calculating the percentage change from the closing price on the first trading day of the year compared to the closing price on the last trading day of the year.  In the final, refactored code, this was done using a for loop that aggregated volume and identified the first and last closing prices for the year by cycling the for loop through each row of data for all variables in all arrays. This code is shown below.

>For j = 2 To RowCount
   
>  If Cells(j, 1).Value = tickers(tickerIndex) Then
>    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
>    End If
>  If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
>    tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
>    End If
>  If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
>    tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
>    End If   
>  If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
>    tickerIndex = tickerIndex + 1
>    End If
>Next j

With the refactored code, including the above loop, the run time improved from 1.16 seconds to 0.28 seconds.


## Summary

Refactoring code requires an up front investment in time to rework and improve code. However, this investment up frong results in simple code that is more readable and easier to update in the future, which reduces time required to maintain and update in the future. It also reduces cost and saves time when using large data sets or frequrent transaction rates. In this instance, the refactored code ran faster and will allow Steve to analyze additional stocks without expanding run time beyond practical limits.
