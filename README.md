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
 
 ![array input](https://github.com/mein0819/stock-analysis/blob/main/readMe_Images/array_Input.png)
 
 and then prints the output to the workbook from the information held in the arrays. 

![array output](https://github.com/mein0819/stock-analysis/blob/main/readMe_Images/print_Array_Info.png)

Holding the information in arrays allows added efficiency by storing the data once in one loop and printing in another loop once the arrays are full, 
rather than storing the information in a variable, printing, and then overwriting the variable for the next iteration. The physical evidence shows in the speed 
test that is performed. 

Original Code: &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  Refactored Code:

![2017original](https://github.com/mein0819/stock-analysis/blob/main/readMe_Images/stock_Analysis_2017_Time.png) &nbsp; &nbsp; &nbsp; &nbsp;![2017refactored](https://github.com/mein0819/stock-analysis/blob/main/readMe_Images/stock_Analysis_Refactored_2017_Time.png)


Original Code: &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  Refactored Code:

![2018original](https://github.com/mein0819/stock-analysis/blob/main/readMe_Images/stock_Analysis_2018_Time.png) &nbsp; &nbsp; &nbsp; &nbsp;![2019refactored](https://github.com/mein0819/stock-analysis/blob/main/readMe_Images/stock_Analysis_Refactored_2018_Time.png)

The refactored code increased the speed of the program by a little over 8x which is pretty significant when analyzing a data set as large as the stock market.

## Summary
### Pros and Cons for Refactoring Code
Refactoring is beneficial to making code more organized and easier to understand, making it more efficient, locating bugs, and keeping the code adaptive and up to date. Problems with refactoring may occur if the designer of the new code isn't familiar with the system and background of the software, which can lead to new bugs, disrupting a stable system. Time and money also are a factor in whether refactoring is actually beneficial as the goals of the project always have to be kept in mind.

### Refactoring Code for this Project
Refactoring the code to conduct this analysis increased the efficiency and improved the organization of the script. The results in the speed test prove this. As far
as any disadvantages for refactoring, I don't believe this program is too large or too complicated that any future developer would have an issue in any design 
improvement.
