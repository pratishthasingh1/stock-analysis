# Stock Analysis
There are many green energy stocks to invest in. An excel file contains all of these stocks for the years 2017 and 2018. Microsoft Excel's VBA was used to determine which of these stocks are worth investing in. The goal was to determine this in a more efficient way by using VBA instead of manually analyzing. 

## Overview
A tickerIndex was set to zero before looping ove rthe rows. Tickers, tickerVolume, tickerStartingPrices, and tickerEndingPrices were put in an array. tickerIndex was used here to access the stock ticker index for these arrays. 
While the script loops through the stock data, it was reading and storing values from each row. The cells were then formatted to show green if the values are greater than 0, and red if the values are less than 0. After using VBA, one can see that the outputs for the 2017 and 2018 DQanalysis stock match the all stock analysis sheet output; however, the result was just found in a faster time.

## Results
The results in the all stock analysis is below: 
![all stock]()

The results of the 2017 analysis:
![2017 stock]()

The results of the 2018 analysis:
![2018 stock]()

## Summary
The advantages of refractoring code in general is it's faster and more organized. The program itself can help make debugging easier. Using whitespaces between code allows commenting which is beneficial to those who want to see how the analysis was done. 
However, there are some disadvantages to refactoring code. If an application is too large for the existing code, it may make refactoring more difficult. However, refactoring can slow down development.