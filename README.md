
# Excel VBA Stock Analysis Code Refactoring

## Overview of Project
Given a years listing of all stocks on an Excel sheet, it is desired to scan through the list and find performances of our stocks of interest. The first VBA coding of this used a nested for loop.  The outer loop incremented through the list of stocks of interest while the inner loop stepped through each record of stocks.  Due to performance limitations it was desired to refactor this code.

## Analysis 

By updating the code to only step through the stock data one time and using an array to store stock performance the following improved times were achieved.

![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/VBA_Challenge_2017.PNG)


![alt text](https://github.com/jj2773/Excel-VBA-Stock-Analysis/blob/main/VBA_Challenge_2018.PNG)


## Summary

There are advantages to the new code of speed, but the disadvantage is ......