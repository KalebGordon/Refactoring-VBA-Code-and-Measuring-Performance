# Refactoring-VBA-Code-and-Measuring-Performance

## Overview
After preparing an initial workbook for steve to analyze stock data, we were requested to do additional research for his parents. 
Initially, we covered only a handfull of stocks for Steve, now we must analyze the entire market for 2017 and 2018 with a given data set. Here, we refactored our previous
code to be inclusive of the entire market and as fast as possible. 

## Results
Our results are dependent on two pieces of information. The total daily volume, and the return of a stock. As a quick rundown, the total daily volume of a stock
is the number of shares that are traded for each day in our time range. The return of a stock is the change in value of the stock in our time range. The return is 
represented by a percent that shows how much the stock increased or decreased. Our overarching goal was to determine which stocks would be better to invest in.

Overall, our analysis covered 12 stocks over two years, 2017 and 2018. We determined that 2017's average return was 67.3% compared to 2018's
-8.5% average using Excel's AVERAGE(Range) function. In addition, 2017 stocks' combined total daily volume was 3,166,639,100 while 2018 stocks' was 3,306,038,200. Overall, 2017 was a far better year for our stocks,
however we should base our conclusions on 2018's results as it is the closest to the present stock market. This was done with Excel's SUM(Range) function. 

The top three highest traded stocks of 2018 were ENPH at 607,473,500 trades, SPWR at 538,024,300 trades, and RUN at 502,757,100 trades. Our only two stocks with a 
positive return in the data set are also among the top three highest traded. RUN had a return of 84.0% and ENPH had a return of 81.9%. Based on this analysis, the best 
stocks to invest in would be ENPH or RUN, with ENPH having a higher total daily volume and RUN having a higher return. 

In order to pull this data from the original data set, we had to loop through each row on our spreadsheet and extrapolate values to put it into the new spreadsheet. For the total daily volume, we made a variable, set it to zero at the beginning of each iteration, then added each value in the total daily volume cell for each stock. Afterwards, we checked if the row beforehand was a different stock. Then, we would set the starting price of that stock if it fulfilled that parameter. The same method was used for the ending price, but we instead checked if the next row was a different stock. 

Our coding that we used is within the VBA in the excel sheet under the appropriate comments. 


2017 Analysis             |  2018 Analysis
:-------------------------:|:-------------------------:
![image](https://github.com/KalebGordon/Refactoring-VBA-Code-and-Measuring-Performance/blob/main/Resources/VBA_Challenge_2017.PNG)  |  ![image](https://github.com/KalebGordon/Refactoring-VBA-Code-and-Measuring-Performance/blob/main/Resources/VBA_Challenge_2018.PNG)


For our code's runtime, we were able to reach a time of 0.130127 seconds for 2017. For 2018, we had a runtime of 0.1140137 seconds. 

## Advantages and Disadvantages of Refactoring
Refactoring can have a lot of advantages and disadvantages. Refactoring can greatly reduce the runtime of our code. In a case where we would need to iterate through tens of thousands of rows while pulling multiple objects from each row, it can exponentially reduce the runtime. It can also make our code easier to understand. If our original code is dozens of if then statements or for loops, we may be able to condense them into less lopps or if then statements. This would make our blocks much easier to read, especially when we hand the project over to other editors. 

This, however, is typically only beneficial if the amount of work done to refactor is less than the benefit of faster code. If something will only need to be run once for an analysis, for example, it may not be very beneficial to refactor. It is greatly beneficial when the code needs to be run multiple times for different purposes. If a code will be run thousands of times (or automatically based on an input) the amount of time saved from refactoring could be very beneficial. However, if the program only needs to run once, and it takes hours to refactor our work, it may not be very helpful to shave off a few seconds from our runtime.  


