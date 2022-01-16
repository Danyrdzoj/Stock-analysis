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

### Differences in Time Execution
In "All Stocks Analysis Refactored", Goal was to refactor the code to increase the efficiency of the original code.
With the help of setting timer, it shows how much time it required to run the process or code.

![Image](VBA_Challenge_FirstResults.png?raw=true)
![Image](VBA_Challenge_Results2017Time.png?raw=true)
![Image](VBA_Challenge_Results2018Time.png?raw=true)

###Results of the Stock Development in 2017 and 2018
Comparing the 2017 and the 2018 stocks, the difference in the total daily volume between the two years that resulted in less than a $100,000,000 in increased volume was not enough to generate a positive 2018 return percentage. The tickers ENPH and RUN had would have been considered good investments due to the positive returns in 2018, and both of the tickers had increases greater than $200,000,000 over the 2017 total daily volumes

![Image](VBA_Challenge_Results2017.png?raw=true)
![Image](VBA_Chellenge_Results2018.png?raw=true)

### Differences in the Code
The original code contained a nested for loop. Included in that outer loop, it also output the data for the current data, which resulted in a significant number of iterations.

## Conclusions
### Advantages
Refactoring can optimize the code efficiency like we have seen in the challenge, also it can help figure out and debug the VBA code. In refactoring, duplicated subroutines, unnecessary loops, redundant statements or simply a faulty code can be removed and debugged. Another advantage when refactoring codes is that looking at the code from a fresh perspective might provide a chance to improve the code.
### Disadvantages
Different programers have different approaches to write codes, and programmers have different logics sometimes this requires testing of the different approaches taken to determine which code is better which can consume effort and time. Refactoring an existing code might possibly introduce other bugs into the code and results in further issues. It can sometimes be challenging to determine weather to start a new code or refactor and existing one.
### 2. How do these pros and cons apply to refactoring the original VBA script?
We have seen that the refactored code reduced the number of loop required thus decreased the processing memory needed for processing the stocks data. This optimized the VBA code which can be observed as the runtime between the original and refactored code significantly decreased. To refactor the VBA code testing must be done to determine if the efficiency has improved or not.
