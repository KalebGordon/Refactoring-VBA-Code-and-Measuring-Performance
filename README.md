# Refactoring-VBA-Code-and-Measuring-Performance

## Overview
After preparing an initial workbook for steve to analyze stock data, we were requested to do additional research for his parents. 
Initially, we covered only a handfull of stocks for Steve, now we must analyze the entire market for 2017 and 2018 with a given data set. Here, we refactored our previous
code to be inclusive of the entire market and as fast as possible. 

## Results
Our results are dependent on two pieces of information. The total daily volume, and the return of a stock. As a quick rundown, the total daily volume of a stock
is the number of shares that are traded for each day in our time range. The return of a stock is the change in value of the stock in our time range. The return is 
represented by a percent that shows how much the stock increased or decreased. 

Overall, our analysis covered 12 stocks over two years, 2017 and 2018. We determined that 2017's average return was 67.3% compared to 2018's
-8.5% average using Excel's AVERAGE(Range) function. In addition, 2017 stocks' combined total daily volume was 3,166,639,100 while 2018 stocks' was 3,306,038,200. Overall, 2017 was a far better year for our stocks,
however we should base our conclusions on 2018's results as it is the closest to the present stock market. This was done with Excel's SUM(Range) function. 

The top three highest traded stocks of 2018 were ENPH at 607,473,500 trades, SPWR at 538,024,300 trades, and RUN at 502,757,100 trades. Our only two stocks with a 
positive return in the data set are also among the top three highest traded. RUN had a return of 84.0% and ENPH had a return of 81.9%. Based on this analysis, the best 
stocks to invest in would be ENPH or RUN, with ENPH having a higher total daily volume and RUN having a higher return. 

![image](Refactoring-VBA-Code-and-Measuring-Performance/Resources/VBA_Challenge_2017.png)

![image](Refactoring-VBA-Code-and-Measuring-Performance/Resources/VBA_Challenge_2018.PNG)

For our code's runtime, we were able to reach a time of 0.130127 seconds for 2017. For 2018, we had a runtime of 0.1140137 seconds. 

## Advantages and Disadvantages of Refactoring
