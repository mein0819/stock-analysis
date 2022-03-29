# **Stock Analysis with VBA**
## Overview of Project
### Purpose
The purpose of this project is to help the client analyze green energy stocks through their total daily volume and percentage
of yearly return over 2017 and 2018. In an attempt to analyze a larger data set of stocks the original analysis code has been 
refactored to improve the efficiency and speed of the analysis. 

## Analysis and Results
### Analysis of Stock Performance 2017 and 2018
The stocks for 2017 had a much better performance, averaging a 67.3% return on a total daily volume average of 263,886,592 shares. Compartavily,
their performance for 2018 took a huge dip with the average return dropping all the way to -8.5% on a total daily average of 275,503,183 shares.
Out of the 12 green energy stocks analyzed, only two had a positive return for 2018, both performing quite well over the year. 

![2018 table](https://github.com/mein0819/stock-analysis/blob/main/readMe_Images/2018_chart.png)

Overall though, it may be wise for the client to advise his parents to expand their search of green energy stocks that show a better return over time. 
Fortunately, the next section will show that the refactored code for this analysis improves efficiency and speed to run on a much larger data set.

### Refactored Code vs. Original Code Speed Differences
The code from the All Stocks Analysis script was refactored to store the necessary data from the workbook in arrays. A unique array was created to hold
each of the total volume, starting price and ending price for each "ticker" that was stored in a ticker array.

![data arrays](https://github.com/mein0819/stock-analysis/blob/main/readMe_Images/arrays_Refactored.png)
 The code loops through the rows of the selected workbook using the variable tickerIndex as the iterator and stores the appropriate information from selected 
 cells into the appropriate array through a series of conditions...
 
 
