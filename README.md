# VBA-challenge

This project was to develop a macro to analyze an array of daily stock trading information and compute annual trading results for multiple years of data.


Design concept:

My design is to treat the data as an unsorted array and use nested for loops to scan the entire array for all instances of each individual value. The script will pull the trading volumes and add them to the annual total during this loop. Since the daily records in the array do not all start and end on the first and last of the year, there will be a second for loop to walk through all of the value to determine their first and last dates. 



Calculated Values:

Annual change: The Script determines the earliest date and latest date any trading was conducted for each ticker symbol by comparing each rows date value to the last value saved for each date type. If the new date is lower than the earliest date variable, it replaces that value. If it is greater than the latest date variable, it replaces that value. If not, it is skipped. When the date replacement occurs, the associated opening value or closing value is also updated in a variable for opening value for the earliest date, or closing value for the latest date. The data set contains some zero values in the opening date. This can cause a divide by zero fault when calulating percentage change so zero values are treated as an empty date and skipped. When the next ticker value is detected, The annual change is calculated simply by subtracting the opening value variable from the closing value variable. 

Percent change: uses the annual change computation and divides it by the opening value. If the opening value is still zero, then the annual change is also zero, so an if, then, else statement is used to set the percent change as O%.

The annual trading volume is simply added as each row is interogated.