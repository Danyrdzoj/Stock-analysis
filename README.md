# Stock-analysis
Steve´s stock analysis by using VBA

## Overview of the Project 
We will help Steve to analyze stock market by:
Creating a Visual Basic macro that can output values we need for our analysis, we need a code that can analyze, read and change cell values, and also we need that the formatting of the spreadsheets can be done using a macro Prograaming structure.
We can get the analysis by creating loops and conditionals to direct logic flow 
We are analyzing the yearly stock data to compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

## Purpose of the Project 
We are analyzing the yearly stock data to compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
For this second challenge, we’ll edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataser. Then, we’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, we just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

## Results of the Project 
### Creating Arrays 
After creating the three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices, the tickerIndex was used to access the stock ticker indices. For this a loop was created to initialize the tickerVolumes to zero and if the following did not match to the before then we increased the tickerIndex. This loop also loops over all the rows in the spreadsheet.

![Image](VBA_Challenge_Macro.png?raw=true)

##Differences with the first Code
In "All Stocks Analysis Refactored", Goal was to refactor the code to increase the efficiency of the original code.
With the help of setting timer, it shows how much time it required to run the process or code.

![Image](VBA_Challenge_FirstResults.png?raw=true)
![Image](VBA_Challenge_Results2017Time.png?raw=true)
![Image](VBA_Challenge_Results2018Time.png?raw=true)
