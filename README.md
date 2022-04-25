# Stock Analysis with Excel and VBA
## Overview of Project

### Purpose and Background

Using the stock data from Steve [(download here)](https://github.com/Bulzeye89/stock-analysis/blob/main/VBA_Challenge.xlsm) for 12 stocks for the years of 2017 and 2018, I was able to create a VBA script while working through module 2.  The macro from this VBA script allowed Steve, with a click of a button, to easily analyze the performance of the 12 stocks in the research for his parents.  This analysis can be seen below.  
<p float="left">
<img src="https://github.com/Bulzeye89/stock-analysis/blob/main/Resources/2017%20Stock%20performance.png" width=30% height=30%>
<img src="https://github.com/Bulzeye89/stock-analysis/blob/main/Resources/2018%20Stock%20performance.png" width=30% height=30%>
</p>

He has now inquired if we can expand the data set to include the entire stock market over the last few years.  While our original code ran fairly quickly, with such a larger data set, the code could get bogged down.  Our first step in accomplishing this goal for him is to dive back into our code and refactor it so that it can run more efficiently. 


## Results

In our original code, the macro had to run through the data 12 times, once for each ticker.  The biggest difference in our refactored code is we took this:

    'Output the data for the current ticker    
        Worksheets("All Stocks Analysis").Activate    
          Cells(4 + I, 1).Value = ticker
          Cells(4 + I, 2).Value = totalVolume
          Cells(4 + I, 3).Value = endingPrice / startingPrice - 1
      
 and created 3 output arrays with this code
 
    'Create three output array
        Dim tickerVolumes(12) As Long    
        Dim tickerStartingprices(12) As Single    
        Dim tickerEndingprices(12) As Single
    
Doing this, allowed us to insert code inside our loop to write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker

    'Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
        
and more efficiently create the output for the Ticker, Total Daily Volume, and Return using the code below.  

    'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For k = 0 To 11
          Worksheets("All Stocks Analysis").Activate             
          Cells(4 + k, 1).Value = tickers(k)    
          Cells(4 + k, 2).Value = tickerVolumes(k)    
          Cells(4 + k, 3).Value = tickerEndingprices(k) / tickerStartingprices(k) - 1
          
The increase in efficiency can be seen in the timer function in the MsgBox that is in both the original code and the refactored code that can be seen in images below. Both 2017 and 2018 stock analysis using the original code took roughly 0.27 seconds while both years using the refactored code took roughly .07 seconds.  


<p float="left">
<img src="https://github.com/Bulzeye89/stock-analysis/blob/main/Resources/Run%20Time%202017%20Original%20code%20.png" width=20% height=20%>
<img src="https://github.com/Bulzeye89/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png" width=20% height=20%>
<img src="https://github.com/Bulzeye89/stock-analysis/blob/main/Resources/Run%20time%202018%20original%20code.png" width=20% height=20%>
<img src="https://github.com/Bulzeye89/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png" width=20% height=20%>
</p>  

## Summary

In general, the biggest advantage of refactoring code is that it can make it more efficient, cleaner, less complex and thus easier to maintian.  The biggest downside is that it can be very time consuming, especially if you are refactoring code that isn't your own or that you haven't worked with in some time.  While I was working through refactoring this code for the stock analysis, it was very time consuming to debug my logic and errors as I'm still learning the syntax to essentially save microseconds to run a code.  It was very good learning process to show that there are different paths to get to the same point and that there is an art in structuring the logic of code.  In addition, the original analysis had three different buttons in excel to basically run the 3 different macros.  With the refactored code, I was able to combine all three buttons/macros into one as illustrated below.  

<img src="https://github.com/Bulzeye89/stock-analysis/blob/main/Resources/Stock%20analysis%20buttons.png" width=50% height=50%>
