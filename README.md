# VBA-challenge

This project was to develop a macro to analyze an array of daily stock trading information and compute annual trading results for multiple years of data.
In all there were over 2000 different ticker symbols with over 800,000 daily records in the multiyear data set.

Design concept:

My design is to treat the data as an unsorted array and use nested for loops to scan the entire array for all instances of each individual value. The script will pull the trading volumes and add them to the annual total during this loop. Since the daily records in the array do not all start and end on the first and last of the year, there will be a second for loop to walk through all of the value to determine their first and last dates.

Having problems getting program to run completely with multi-year data set. Takes too long to run. 



Calculated Values:

Annual change: The Script determines the earliest date and latest date any trading was conducted for each ticker symbol by comparing each rows date value to the last value saved for each date type. If the new date is lower than the earliest date variable, it replaces that value. If it is greater than the latest date variable, it replaces that value. If not, it is skipped. When the date replacement occurs, the associated opening value or closing value is also updated in a variable for opening value for the earliest date, or closing value for the latest date. The data set contains some zero values in the opening date. This can cause a divide by zero fault when calulating percentage change so zero values are treated as an empty date and skipped. When the next ticker value is detected, The annual change is calculated simply by subtracting the opening value variable from the closing value variable. 

Percent change: uses the annual change computation and divides it by the opening value. If the opening value is still zero, then the annual change is also zero, so an if, then, else statement is used to set the percent change as O%.

The annual trading volume is simply added as each row is interogated.

Next the program creates a new formated area and looks through the newly created array of annual data to find the best and worst performing stocks by percent change simply by running a set of commands to return the max and min values from the percent change column. And doing the same to the trading volume column to find the stock with the highest trade volume. The commands also return the row information for each value, which is used to obtain the associated ticker symbol information.


Conditional Formatting: 

I used the 'record macro' function to manually set up the conditional foratting for the data columns, then edited the code to make it run in the final Macro.

Entire Workbook functionality:

I used a 'For Each' Loop at the beginning of the code to iterate the code through all of the worksheets in the workbook. I used the variable page assigned as a worksheet. This also required editing the entire code to add the 'page.' modifier to the begging of every cell, range, and column refernce in the code, without that, all updates are only made to the first worksheet.


Reset Macro: 

There is also a macro that can be run to erase all of the data and conditional formatting in order to quickly reset the page to its original state in order to observe follow on test runs
