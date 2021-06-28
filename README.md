# **An analysis of the Stock Market**

## Overview of Project

An analysis of the Stock Market is a project aimed to help our client strategize his investments by providing him with an automated analysis of a portfolio of stocks over a period of time. Utilizing Excel and Visual Basic Applications (VBA), we designed a code to analyze the _total daily volume_ and _yearly return_ for 12 stocks over 2017 and 2018. The result offers the opportunity to carry the analysis in less than two seconds, with a click of a button! 

### Results

This project equipped the client with a tool to carry out a comprehensive comparative report of each of the 12 stocks' performance between 2017 and 2018. The end result is easy to use, understand and has reduced the time to run the analysis from manual to automated: [VBA_Challenge](https://github.com/chocoplace/stock-analysis/blob/main/VBA_Challenge.xlsm). 

*Baseline of the project:

Initially the client was interested only in learning how actively the stock “DQ” was traded in 2018 and its yearly return, information needed to evaluate future investments. 

The analysis began with the objective of identifying the behavior of a specific stock but later we expanded, per the client request, to include the entire stock market over the last few years. This request requires us to refactor or original code to provide an even faster way to analyze the data. 

*Elements of the code*: We wrote a code to output the information from the dataset [Green_Stocks](https://github.com/chocoplace/stock-analysis/blob/main/green_stocks.xlsm) with three key components: Ticker, Total Daily Volume and Return. The code includes different elements and techniques such as: variables, nested loops, conditionals, formatting, among others. The refactored code can be revise here: [VBA_Challenge_Final](https://github.com/chocoplace/stock-analysis/blob/main/VBA_Challenge_Final.vbs)

- *_Ticker_*: The list of the 12 stocks analyzed on this project to be able to carry out a comprehensive analysis and make a well-informed decision. Includes: AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP and VSLR.  

- *_Total Daily Volume_*: The sum of all the daily tradings per stock used to identify the value of the stock. The behavior (how often the stock is traded) of the stock reflects the value of the stock.

- *_Return_*: Used to measure the yearly performance of the stock, is the percentage increase or decrease in price from the beginning of the year to the end of the year. It applies for example: “if the client invested in DQ at the beginning of the year and never sold, the yearly return is how much the investment grew or shrunk by the end of the year”. 

- *_Years_*: The analysis was conducted with the information generated over 2017 and 2018. 

*Findings:

- According to the analysis results, 2017 was a great year for the portfolio of stocks, reporting 11 of 12 stocks with positive returns from the beginning of the year.
- DQ delivered the highest return with a 199.45% of increase and TERP was the only stock that reported losses with a -7.21% decrease. 
- The top three stocks of 2017 are DQ with a 199.45% increase from the beginning of the year or $35,796,200.00 of stock value; followed by SEDG with a yearly return increase of 184.47% or $206,885,200.00 of value; and ENPH with an increase of 129.52% and a value of $221,772,100.00. 
- The analysis is conducted in 1.3125 seconds. 

With refactored code: 

![VBA_Challenge_2017](https://github.com/chocoplace/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

With original code: 

![Green_Stocks_2017](https://github.com/chocoplace/stock-analysis/blob/main/Resources/Green_Stocks_2017.png)

- According to the analysis results, 2018 was a difficult year for the portfolio of stocks. Only two of the 12 stocks reported positive returns from the beginning of the year.
- Among the stocks that delivered positive returns, we can highlight RUN with the highest return with an 83.95% increase and a $607,473,500.00 value. 
- Among the stocks with the lowest returns are DQ with a -62.60% decrease, JKS with a -60.53% of decrease, and SPWR with a -44.59% of decrease.
- The analysis is conducted in 1.2343 seconds.

With refactored code:

![VBA_Challenge_2018](https://github.com/chocoplace/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

With original code:

![Green_Stocks_2018](https://github.com/chocoplace/stock-analysis/blob/main/Resources/Green_Stocks_2018.png)

We can conclude the following:  

- The initial interest of our client was to invest in DQ, however the analysis can predict a trend of the instability of the stock’s performance. 
- The analysis shows ENPH and RUN as the two stocks that have a positive and steady performance. The client can evaluate the option to invest. 
- Expanding the analysis from one to twelve stocks broadened the scope of the analysis and offered us the opportunity to improve our code, making it better and faster.

#### Summary

*Advantages and disadvantages of refactoring code in general*

- Advantages: One of the benefits of refactoring code, in general, is to have the opportunity to improve the performance by tidying up the script. For a beginner is an opportunity to identify patterns or re-usable code for future projects and apply the “Don't Repeat Yourself” rule, among other learned techniques. 

- Disadvantages: One of the limitations of refactoring code, in general, is the relation between the amount of time invested in improving the script and the functionality, investing time in improving the code does not ensure an improvement inefficiency. For beginners, refactoring code can be challenging even more if the code was created by another programmer.


*Advantages and disadvantages of the original and refactored VBA script*

- Advantages: As a student of data analysis and visualization, the major advantage or improvement I experienced while refactoring the VBA script was to practice and understand at a deeper level each function of the code. In particular with the original code, the script for formatting was in a different subroutine and was challenging connecting both (at the end I created a button for format), and by refactoring the code I was able to include the formatting on the same script and have all the analysis and formatting in ONE click. See picture: 

![Final_buttons](https://github.com/chocoplace/stock-analysis/blob/main/Resources/Final_buttons.png)

- Disadvantages: One of the limitations I see while working on refactoring the code was having to implement new actions on the code while keeping the pattern. For me, it presented a challenge because I haven't perfected the skill to code and the exercise forced me to go over my own code. If this was a challenge with a code I created, I can imagine the challenge of working with a preexisting code. Another disadvantage was determining if the refactored code made a minimum impact on the efficiency of the VBA script, the new code runs slightly faster than the original. 

End



