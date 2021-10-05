# Stock-Analysis Overview of Project
## Purpose

The purpose of this project was to editing a Microsoft Excel VBA code to process certain stock information in the year 2017 and 2018 and determine whether or not the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this project is to increase the efficiency and speed of the origin 2l code.

## The Data

The data that is presented includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the tickers, the total daily volume, and the return on each stock.

## Results
We were give a code that walked us through certain procedures that includeded:
  1a) Create a ticker Index
  1b) Create three output arrays
  2a) Create a for loop to initialize the tickerVolumes to zero.
  If the next row’s ticker doesn’t match, increase the tickerIndex.
  2b) Loop over all the rows in the spreadsheet.
  3a) Increase volume for current ticker
  3b) Check if the current row is the first row with the selected tickerIndex.
  3c) check if the current row is the last row with the selected ticker
  3d) Increase the tickerIndex.
  '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
  These are the steps we used to along with creating input box, chart headers, ticker array, and to activate the appropriate worksheet. These process helps us to set the structure for the refactoring

# Summary
## Advantages and disadvantages of refactoring code in general 
The process of refactoring the code allows us to improve the code in order to make it easier to work with. This can be done without changing or adding to its external behavior. The benefits of this is that it can save time, reduced complexity and scalability. The disadvantages are that it can be time consuming and potentially the reduced code may not imporve the run time. 

## Advantages and disadvantages of the original and refactored VBA script 
fter refactoring our dataset a Advantages was a shorter run time for both 2017 & 2018 that can be seen in the screen shots below. The original analysis took over one second to run, whereas our new analysis has redeuced the time to less than that.


![VBA_Challenge_2017](https://user-images.githubusercontent.com/31675832/135778379-a79dd803-ae9e-4b5b-ba63-20ab90e76d16.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/31675832/135778416-dc26b09b-c855-48d2-89e0-e9638faf8113.PNG)
