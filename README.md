# stock-analysis
* VBA analysis of green energy stocks

# Stocks Analysis with Excel VBA

## Overview of Project

### Purpose

A VBA macro function was created to automate an analysis of green stocks for the client, Steve. Although the analysis works for a small sample dataset, it may not be efficient when analyzing large datasets. The goal here is to refactor the VBA macro by editing and cleaning the VBA script in order to decrease its run time and improve performance when performing on larger datasets. 


## Results

### Stock Performance

Almost all stocks grew positively during 2017, except for the ticker: TERP which decreased by 7.2% over the year (Table 1). 

![img_1](https://github.com/jmasurovsky/stock-analysis/blob/master/Resources/VBA_Challenge_AllStocks_2017.png)

Table 1. 2017 stock performance

In contrast, almost all stocks decreased in 2018, except for the stock tickers:
ENPH and RUN (Table 2). 

![img_2](https://github.com/jmasurovsky/stock-analysis/blob/master/Resources/VBA_Challenge_AllStocks_2018.png)

Table 2. 2018 stock performance

Even though almost all stocks dropped in 2018, most of their growth in 2017 was greater than the following year, making their 2 year return as positive growth. For example, VSLR stock grew 50% in 2017 and dropped 3.5% in 2018, therefore leading to a 46.5% increase in their stock over 2 years. -DQ?

### Code Execution Times

Before refactoring the VBA script, the executions times for the “AllStocksAnalysis” macro ran for about 0.62 seconds for the year 2017 (Figure 1), and about 0.63 seconds for the year 2018 (Figure 2).

![img_3](https://github.com/jmasurovsky/stock-analysis/blob/master/Resources/VBA_Challenge_2017_%20NotRefactored.png)

Figure 1. Original VBA script run time for 2017


![img_4](https://github.com/jmasurovsky/stock-analysis/blob/master/Resources/VBA_Challenge_2018_NotRefactored.png)

Figure 2. Original VBA script run time for 2018

After refactoring the VBA script, the execution times for the “AllStocksAnalysisRefactored” macro ran for 0.11 seconds for the year 2017 (Figure 3), and about 0.13 seconds for the year 2018 (Figure 4). Overall, the refactored VBA script ran faster for both years.

![img_5](https://github.com/jmasurovsky/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

Figure 3. Refactored VBA script run time for 2017


![img_6](https://github.com/jmasurovsky/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

Figure 4. Refactored VBA scirpt run time for 2018


-Add images to code and discuss difference of nesting for loops (they can lead to longer execution times) vs separating them. Creating an index?

### Summary

## Pros and Cons of Refactoring

Refactoring code is the process of taking previously written code and editing it to make it more readable, clean, and understandable thus making it more efficient without changing its function and output. Refactoring is typically performed because it helps perform tasks on larger datasets, applications, and easier to find errors in the code. It is important to know when to finish refactoring code in order to focus on completing the task at hand, because refactoring code can be time consuming.

Refactoring the VBA script increased its performance by decreasing run time and making the code more readable to debug. The disadvantage was having to edit and add more code in the refactored VBA script compared to the original, such as creating arrays, an index to be increased and referenced throughout the code.

