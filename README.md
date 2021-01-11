# Green-Stocks_analysis
Challenge 2 - Performing analysis on Green Stock data with VBA

## Overview of Project

In this project, a analysing of Green Stock data was provided using VBA.  By looping through the data of the Year of 2017 and 2018, total daily volume and return for each ticher was culculated.  After refactoring the code, the script is running more efficiently than before (After: 7+s vs Before: 9+s). To achieve the refactoring, unensessary steps were deleted, logic of the code structure was updated and functionalities was added.

## Analysis and Challenges

### The Overview of Analysis

This analysis is based on the pre-analysis of the [green_stocks.xlsm](/green_stocks.xlsm) workbooks.  The code is working with high volume of stocks for the entire stock market.  Nested loops was created inside the code to ensure all data in selected spearsheets could be go through.

### The Overview of Analysis - Theater Outcomes by Launch Date

The purpose of **"Theater Outcomes by Launch Date"** analysis is to present the counts of successful, failed and canceled resords for theater by monthly based statistics of launched dates. Louise could intuitively view the trends of outcomes according the month launched with line chart.

![Theater_Outcomes_vs_Launch](Resources/Theater_Outcomes_vs_Launch.png)

### The Overview of Analysis - Outcomes Based on Goals

The Purpose of **"Outcomes Based on Goals"** visualized the data of posibility of successful, failed or canceled for plays under each range of goals.  This chart determins the posibility of outcomes in different range of goal for the plays, in order to predict the posibile outcome for Fever based on the fundraising goal amount.

![Outcomes_vs_Goals](Resources/Outcomes_vs_Goals.png)

More details could be found in [Kickstarter_Challenge](/Kickstarter_Challenge.xlsx.zip)

### Challenges

1. When I working on Deliverable 3, I found the relative links are too long when create relative links for pictures, and it would he an error once the files relocate on computer.  To solve this problem, I uploaded the PNG files to the cloud firstly, then use the source uploaded to disaplay in the README.md file.

2. I got the wrong result of the chart at the beginning when I did Deliverable 2 at the beginning.  Then I go back to the requirments step by step more carefully and found that I forgot to set criteria for subcategory. I went back to reset the fomular and finally got the right result.

## Results
- What are two conclusions you can draw about the Theater Outcomes by Launch Date?
  * The numbers of successful outcomes is the highest over the year in May. For Maximizing the posibilyty of succeed, Louise could launch the play in May.
  * There is a low posibility to fail in November. If Louise is a risk avoider, November would be the best period for launching.

- What can you conclude about the Outcomes based on Goals?
  * There is a high posibility to fail if the fundraising goal is between 25000 and 29999.
  * There is a low possibility (almost 0) that the plays are cancled.
  
- What are some limitations of this dataset?
  * There are file size limitations in Excel. If there is a large number of data in this dataset, Excel might not able to handle.
  * In the worksheet "Outcomes Based on Goals", there are some data of Goals with large diference with the most of the numbers. These numbers might cause error when making desision based on the calculating results with these factors.
  
- What are some other possible tables and/or graphs that we could create?
  * We could creat a table for culculation the mean, mode, median and standard diviation of the goal to visualize degree of dispersion for the fundraising goals. Meanwhile quartiles could be calculated for culculating IQR.
