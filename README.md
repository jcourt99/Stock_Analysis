# Green Stock Analysis with VBA

## Overview of Project

### Purpose
The purpose of this project is to analyze data from the stock market with code that executes quickly. The code will look at the data from all stocks in one loop rather than looking at each stock individually and populating the results before moving on to the next stock. By refactoring the code, the execution time to analyze the data will be much quicker.

## Results

### Analysis of Refactored Code for Project
The original code looked at one ticker at a time before populating the data for the tickers. This code worked fine for the 12 tickers in the original analysis, however, once more tickers are added to the project, the execution time will increase greatly as the code moves through one ticker at a time. The execution times for years 2017 and 2018 are captured in the following screenshots as 1.34 seconds and 1.23 seconds, respectively. ![This is the 2017 time screenshot](https://github.com/jcourt99/stock-analysis/blob/main/Pre_VBA_Challenge_2017.png) and ![This is the 2018 time screenshot](https://github.com/jcourt99/stock-analysis/blob/main/VBA_Challenge_2018.png)

By refactoring the code for the stock market analysis, the execution times decreased. The refactored code uses a For loop to look at all of the tickers before populating the output onto the worksheet. The code  looks for a different ticker name, just as in the original code, but then increases the tickerIndex by one in order to go through all 12 tickers before writing the output. The decreased execution times are shown in the following screenshots as 0.197 seconds for 2017 and 0.191 seconds for 2018. ![This is the refactored 2017 time screenshot](https://github.com/jcourt99/stock-analysis/blob/main/VBA_Challenge_2017.png) and ![This is the refactored 2018 time screenshot](https://github.com/jcourt99/stock-analysis/blob/main/VBA_Challenge_2018.png)


## Summary

### General Advantages and Disadvantages of Refactoring Code
Refactoring code is done to quicken the execution of the code by taking fewer steps to create the same output. It is also an advantage to refactor code to make it easier for other programmers to read or to enhance the logic behind it. 

One of the disadvantages of refactoring code is the time it takes to rewrite it. 

### Advantages and Disadvantages of the Original and Refactored Code for Stock Analysis Project
For this project, the advantage of refactoring the code is the amount of time saved for the execution of the results. The VBA script is easier to understand now that there is one loop for all of the ticker indexes. The original code went through the ticker indexes one at a time which works fine for 12 indexes, however, if there are thousands of indexes, the code will take significantly longer to execute. 




