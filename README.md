# VBA-Challenge
Use VBA scripts to loop through 3 years of stock data and output data analysis

## Background

VBA scripting was used to analyze over 750,000 rows of generated stock market data for 3,000 stocks traded each day of the year spanning the years 2018 - 2020. The output of the script created a summary table for each stock symbol, the yearly change in stock price each year, the percent change in stock price each year and the total volume traded.  The output also denoted which stock experienced the greatest percent increase in price, the greatest percent decrease in price and the greatest volume traded for each of the 3 years.

### Approach

1. Leveraged VBA scripting to loop through all three worksheets within one set of code
2. Conditional arguments were used to perform functions such as setting data, calculating totals and other mathematical functions, and printing data to summary tables
3. "For Loops” where employed to iterate through the data and perform various functions
4. Conditional formatting with colors was used to highlight the positive (green)and negative (red) performers each year

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Used while developing and testing the scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - The final data used for running the scripts, analyze the data and generate the output analysis. 

### References & Resources

* https://excelmacromastery.com/excel-vba-find/
