
# Excel VBA Stock Analysis Code Refactoring

## Overview of Project
Given a full years listing of daily trading values of stocks on an Excel sheet, it is desired to scan through the list and find performances of our stocks of interest. The performance metrics requested are total daily traded volume and the return for the year.  Below is a sample of our raw stock data which is required to be sorted alphebetically by ticker and then chronologically by trading date.

![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/StockDataSample.PNG)


## Analysis 
The first VBA coding for this analysis used a nested for loop.  The outer loop incremented through the list of stocks of interest while the inner loop stepped through each record of stocks.  The nested for loop run times can be seen below for the stock listing for 2017 and 2018.  These runs times exceeded 1 second each and only contained stocks of interest.  If all stocks traded on the market in one year were added to the list the compute time would not be reasonable since just 12 stocks took 1 second.

![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/NestedForLoops_2017.PNG)


![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/NestedForLoops_2018.PNG)


Due to performance limitations it was desired to refactor this code.  By updating the code to only step through the stock data one time and then use an array to store stock performance values of interest the following improved times were achieved.

![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/VBA_Challenge_2017.PNG)


![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/VBA_Challenge_2018.PNG)


## Summary

The new code refactoring is much faster since only one pass is made over the stock data (one for loop).  In the previous code there was 12 passes over the same stock data since there were 12 stocks of interest being checked by the outer for loop.   With this approach if the number of stocks of interest increases then the number of passes, hence compute time, also increases.  This is a major dissadvantage of this approach, but the advantage is simplicity.  Our refactored approach using an array to store data is more complex to debug, but it is much more computationally efficient in the approach.