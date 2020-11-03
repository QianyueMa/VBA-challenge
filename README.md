# VBA-challenge - Stock Market Analysis

## Objectives

To analyze real stock market data using Excel VBA scripting.

![stock Market](Images/stockmarket.jpg)

## Navigation of the repo: Submission

* All the completed files are in the 'VBAStocks' directory.
* The following files are included in the repo:
  * VBA Scripts as separate files: `part1.vbs` contains the VBA script for Part 1, and `part2.vbs` contains the script for Part 2 Challenges. In the modules, the Part 1's macro is named `tickerloop1`, and the Part 2's `tickerloop2`.
  * A screenshot for each year of the results on the `Multi_Year_Stock_Data`.
  * Preview:
[2014](Images/2014.png)
[2015](Images/2015.png)
[2016](Images/2016.png)


## Datasets

* [Testing Data](Resources/alphabetical_testing.xlsx) for the code developing purposes & [Real Stock Data](Resources/Multiple_year_stock_data.xlsx) for the analysis.

- - -

## Tasks

### Part 1

* Create a script that will loop through all the stocks for one year and output the following information.
  * The ticker symbol.
  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The total stock volume of the stock.
 
* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller so the test is faster. The code should run on this file in less than 3-5 minutes.

* Use conditional formatting to highlight positive change in green and negative change in red.

* Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.

* The result looks as follows.
![moderate_solution](Images/moderate_solution.png)

### Part 2: CHALLENGES

1. Return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

![hard_solution](Images/hard_solution.png)

2. Make the appropriate adjustments to the VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once. 

