# Green-Stocks_analysis
Challenge 2 - Performing analysis on Green Stock data with VBA

## Overview of Project

In this project, a analysing of Green Stock data was provided using VBA.  By looping through the data of the Year of 2017 and 2018, total daily volume and return for each ticher was culculated.  After refactoring the code, the script is running more efficiently than before (After: 0.7109375 sec VS Before: 0.7421875 sec). To achieve the refactoring, unensessary steps were deleted, logic of the code structure was updated and functionalities was added.

## Results

### The Overview of Analysis

This analysis is based on the pre-analysis of the [green_stocks.xlsm](/green_stocks.xlsm) workbooks.  The code is working with high volume of stocks for the entire stock market.  Nested loops was created inside the code to ensure all data in selected spearsheets could be go through.  Button was re-assgin to new macro "AllStocksAnalysisRefactored".  By entering the target year, analysis of the entered year will analysed in worksheet "All Stocks Analysis". 

### Main Codes

* Nest Loop Structure:
1        For i = 0 To 11
2             tickerIndex = tickers(i)
3                     For j = 2 To RowCount
4                     Next j
5         Next i

* Nest Loop Code:
    ![VBA_Nest_Loop_Code](Resources/VBA_Nest_Loop_Code.png)
    
* Output Setting Code:
    ![VBA_Output](Resources/VBA_Output.png)
    
More details could be found in [green_stocks.xlsm](/green_stocks.xlsm)

### Challenges

When writing the 

## Summary

As a result, total daily volume and return for 2017 and 2018 are shown as below:

- For the Year of 2017
![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)

- For the Year of 2018
![VBA_Challenge_2018](Resources/VBA_Challenge_2018.png)


- The advantages and disadvantages of refactoring code in general
  * Advantages:
  * Disadvantages:
  
- the advantages and disadvantages of the original and refactored VBA script 
  * Advantages:
  * Disadvantages:
  
