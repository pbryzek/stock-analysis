# Bootcamp: UCB-VIRT-DATA-PT-03-2020-U-B-TTH
### Bootcamp Challenge #2 - 3/22/2020
Bootcamp Challenge 2: Module Stock Analysis with Excel VBA Programming

### Dataset Used
Stock Analysis Dataset 
[StockAnaysisDataSet](https://courses.bootcampspot.com/courses/140/files/35036/download?wrap=1)

### Excel Spreadsheet Generated
[Challenge 2](green_stocks.xlsm)

### Challenge Description
This challenge presented an excel sheet that already had a macro with a hardcoded array of stock symbols that required N iterations of the entire dataset to calculate. The challenge was focused on refactoring the existing code to run only one time across the entire dataset, and store the Volumes, starting prices, ending prices, and ticker symbols for all tickers found in the stock analysis dataset.

Original code had run time coupled with the number of ticker symbols. This new refactoring requires an initial run of the code to find how many symbols there are to initialize the arrays. The function then runs through the set one more time to populate the aforementioned datasets.

Implement buttons to enable dynamic choice to the user for which year to run the analysis on. Refactoring should not change the output of the program and thus I compared the original dataset with the refactored code to ensure quality.

### Challenge Observations
-In 2017 11/12 ticker symbols observed posted positive yearly returns.
-In 2017, the average percentage return among those dozen stocks was 67.32%, with a range of -7.2% - 199.4%.

-In 2018 10/12 ticker symbols observed posted negative yearly returns.
-In 2018, the average percentage return among those dozen stocks was -8.5%, with a range of -62.6% - 84%.
