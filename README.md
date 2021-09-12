
# Excel VBA Stock Analysis Code Refactoring

## Overview of Project
Given a full years listing of daily trading values of stocks on an Excel sheet, it is desired to scan through the list and find performances of our stocks of interest. The performance metrics requested are total daily traded volume and the return for the year.  Below is a sample of our raw stock data which is required to be sorted alphebetically by ticker and then chronologically by trading date.

![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/StockDataSample.PNG)


## Analysis 
The first VBA coding for this analysis used a nested for loop.  The outer loop incremented through the list of stocks of interest while the inner loop stepped through each record of stocks.  The nested for loop run times can be seen below for the stock listing for 2017 and 2018.  These runs times exceeded 1 second each and only contained stocks of interest.  If all stocks traded on the market in one year were added to the list the compute time would not be reasonable since just 12 stocks took 1 second.

![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/NestedForLoops_2017.PNG)


![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/NestedForLoops_2018.PNG)


Due to performance limitations it was desired to refactor this code.  By updating the code to only step through the stock data one time, and then use an array to store stock performance values of interest the following improved times were achieved.  

![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/VBA_Challenge_2017.PNG)


![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/VBA_Challenge_2018.PNG)

Also, a conditional statement was added during the code refactoring to allow the yearly stock listing tables to contain stocks that were not of interest.  The code will just skip over these stocks since they are not in the list of stocks to be analyzed.  


## Summary

It is possible that code refactoring will likely not been seen of much value to stakeholders.  What new features or enhancements are being obtained for the reinvestment of time and money?  If it is only speed, it is possibly seen by the stakeholders as technical debt that should not have existed originally. From the developers viewpoint, code refactoring could provide an opportunity to clean, optimize, and remove potential bugs.

In our code refactoring case, the new code is much faster since only one pass is made over the stock data (one for loop).  In the previous code there were 12 passes over the same stock data since there were 12 stocks of interest being checked by the outer for loop.   With this approach if the number of stocks of interest increases then the number of passes, hence compute time, also increases.  This is a major dissadvantage of this approach, but the advantage is simplicity.  Our refactored approach using an array to store data is more complex to debug and requires separate loops for formatting needs, but it is much more computationally efficient in the approach.  

.