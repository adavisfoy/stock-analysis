# Stock Analysis Code Refactoring
  
## Project Overview

The purpose of this project was to refactor our "All Stocks Analysis" code written in VBA with Excel to handle a larger data set and run more efficiently. While the initial version worked well for a dozen stocks, our customer Steve wants to expand the dataset to analyze the entire stock market over the last few years. Our original code is not ideal for such a large-scale analysis and would likely take a long time to execute. We didn't adhere to the coding mantra "DRY: Don't Repeat YOurself." 

While the output of both versions of code was the same, we implemented reusable code consisting of arrays and nested loops in the refactored version that enabled a complete analysis by looping through all of the data once, rather than numerous times.

## Results

### Stock Performance

In 2017, nearly all stocks realized positive yearly returns ranging from +5.55% to +199.45% (DQ). Only TERP had a negative yearly return of -7.21% in 2017. Based on DQ's performance in 2017, it is understandable that Steve's parents were eager to invest in DQ stock.

With his parents' financial interests in mind, Steve pursued additional analysis revealing that 2018 was not a good year for the stock market. Only 2 stocks had positive returns for the year: ENPH (81.92%) and RUN (83.95%). All others performed poorly for a variety of reasons outlined here: 
(https://www.pbs.org/newshour/economy/making-sense/6-factors-that-fueled-the-stock-market-dive-in-2018)
 
Given that 2018 was not a good year for the stock market, it makes sense that Steve wants to expand his dataset to continue analyzing stocks before his parents make their investment decision. 

### Code Execution Times: Original Code vs. Refactored Code

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

## Project Summary

- Advantages of refactoring code: 
- Disadvantages of refactoring code: 
  - It's Time consuming.
3. How do these pros and cons apply to refactoring the original VBA script?
