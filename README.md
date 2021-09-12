
# Excel VBA Stock Analysis Code Refactoring

## Overview of Project
Given a years listing of all stocks on an Excel sheet, it is desired to scan through the list and find performances of our stocks of interest. The first VBA coding of this used a nested for loop.  The outer loop incremented through the list of stocks of interest while the inner loop stepped through each record of stocks.  Due to performance limitations it was desired to refactor this code.

## Analysis and Challenges


![alt text](https://github.com/jj2773/kickstarter-analysis/blob/main/Theater_Outcomes_vs_Launch.png)

### Analysis of Outcomes Based on Goals
The smaller the goal the more likely the campaign will succeed.  Campaigns keeping their goals under 5,000 had a success rate of around 75%.  While campaigns with goals between 5,000 and 15,000 had a 55% success rate.  Goals chosen higher than 15,000 resulted in 50% failure rates or higher.  The Data samples with goals above 20,000 are not statistically sound due to fewer data points and outliers, so one should not rely on any trends of outcome for data shown in these goal ranges.

![alt text](https://github.com/jj2773/kickstarter-analysis/blob/main/Outcomes_vs_Goals.png)

### Challenges and Difficulties Encountered
Using inequalities in the countif statement for Excel required a double quote due to the equal sign.  Also, the exercise neglected equal to 50,000 in the table which I think is an error in the material.  To correct for this the table should read "45000 to 50000" and "Greater than 50000".

## Results

In conclusion, the most popular and successful launch month is May, and the worst launch month is December. The lower the campaign goal the higher the success rate.  It should be noted that this dataset is primarily made up of United States and Great Britain campaigns in the category of Theater/Plays.  This means that these conclusions are only applicable to those markets.  Statistical tables that include the mean, median, standard deviation, upper quartile, lower quartile, and IQR would give insights into the goal tendencies.  A box and whisker plot would reveal the outlier data points with higher goals that should not be relied upon for conclusions as previously discussed.
#   E x c e l - V B A - S t o c k - A n a l y s i s  
 