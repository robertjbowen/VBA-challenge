# VBA-challenge

This project was to develop a macro to analyze an array of daily stock trading information and compute annual trading results for multiple years of data.
In all there were over 2000 different ticker symbols with over 800,000 daily records in the multiyear data set.

Design concept:

My original Macro design treated the data as an unsorted array and used nested for loops to scan the entire array for all instances of each individual value. This code worked worked on the reduced test data set, but proved to be very long running and could not handle the extremely large data set of the multi-year stock data.

Since the data set was already sorted I refactored the code to use a for loop to step through the array only once and pull out all neccessary values for each computation into a set of variables. The identification of a new ticker symbol triggers the final calculation of the annual totals and a reset of the variables used to make the calculations. The For loop then continues with the new ticker sysmbol until the next symbol is detected. This continues until the entire array is searched and all values are calculated.



Calculated Values:

Annual change: The Script determines the earliest date and latest date any trading was conducted for each ticker symbol by comparing each rows date value to the last value saved for each date type. If the new date is lower than the earliest date variable, it replaces that value. If it is greater than the latest date variable, it replaces that value. If not, it is skipped. When the date replacement occurs, the associated opening value or closing value is also updated in a variable for opening value for the earliest date, or closing value for the latest date. The data set contains some zero values in the opening date. This can cause a divide by zero fault when calulating percentage change so zero values are treated as an empty date and skipped. When the next ticker value is detected, The annual change is calculated simply by subtracting the opening value variable from the closing value variable. 

Percent change: uses the annual change computation and divides it by the opening value. If the opening value is still zero, then the annual change is also zero, so an if, then, else statement is used to set the percent change as O%.

The annual trading volume is simply added as each row is interogated.

After completing the calculations all of the variables are reset to their original values. A row counter is indexed to step to the next row in the ticker symbol column and the new ticker symbol its initial dates, opening and closing values, and first day trading total are added to the list.

The loop then indexes to the next row in the column and the process starts over for the new ticker symbol.

Next the program creates a new formated area and looks through the newly created array of annual data to find the best and worst performing stocks by percent change simply by running a set of commands to return the max and min values from the percent change column. And doing the same to the trading volume column to find the stock with the highest trade volume. The commands also return the row information for each value, which is used to obtain the associated ticker symbol information.


Conditional Formatting: 

I used the 'record macro' function to manually set up the conditional foratting for the data columns, then edited the code to make it run in the final Macro.

Entire Workbook functionality:

I used a 'For Each' Loop at the beginning of the code to iterate the code through all of the worksheets in the workbook. I used the variable page assigned as a worksheet. This also required editing the entire code to add the 'page.' modifier to the begging of every cell, range, and column refernce in the code, without that, all updates are only made to the first worksheet.


Reset Macro: 

There is also a macro that can be run to erase all of the data and conditional formatting in order to quickly reset the page to its original state in order to observe follow on test runs
