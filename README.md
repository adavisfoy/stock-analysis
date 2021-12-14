# Stock Analysis Code Refactoring
  
## Project Overview

The purpose of this project was to refactor our "All Stocks Analysis" code written in VBA with Excel to handle a larger data set and run more efficiently. While the initial version worked well for a dozen stocks, our customer Steve wants to expand the dataset to analyze the entire stock market over the last few years. Our original code is not ideal for such a large-scale analysis and would likely take a long time to execute. 

## Results

### Stock Performance

In 2017, nearly all stocks realized positive yearly returns ranging from +5.55% to +199.45% (DQ). Only TERP had a negative yearly return of -7.21% in 2017. Based on DQ's performance in 2017, it is understandable that Steve's parents were eager to invest in DQ stock.

With his parents' financial interests in mind, Steve pursued additional analysis revealing that 2018 was not a good year for the stock market. Only 2 stocks had positive returns for the year: ENPH (81.92%) and RUN (83.95%). All others performed poorly for a variety of reasons outlined here: 
(https://www.pbs.org/newshour/economy/making-sense/6-factors-that-fueled-the-stock-market-dive-in-2018).
 
Given that 2018 was not a good year for the stock market, it makes sense that Steve wants to expand his dataset to continue analyzing stocks before his parents make their investment decision. 

### Code Execution Times: Original Code vs. Refactored Code

Before we compare execution times for our original code versus refactored code, it is important to note the similarities and differences between them. 
- The output of both versions of code is exactly the same.
- We employed the technique of arrays and nested loops in both versions.
  - The original code utilized only one array for "tickers." We then declared specific variables for "startingPrice" and "endingPrice". We then looped through the data a ticker at a time, which populated the output one row at a time for the current ticker. Because we evaluated each ticker individually, we had to loop through the data 12 times.
  - The refactored code expanded utilization of arrays to four arrays: tickers, x, x, and x. As we looped through the data, we stored the data in the four different arrays and then "dumped" all of the output into our "All Stocks Analysis" worksheet at the end.    

This strategy facilitated a complete analysis of all stocks after only one loop through the data rather than looping through it twelve times (i.e. once per ticker) as we did in the original version.This resulted in greatly improved execution times when we compare the original code vs. the refactored code.  

## Project Summary

- Advantages of refactoring code: 
- Disadvantages of refactoring code: 
  - It's Time consuming.
3. How do these pros and cons apply to refactoring the original VBA script?
