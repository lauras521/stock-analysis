# Stock Analysis for Twelve Tickers in 2017 and 2018
## Overview of Project: 
Steve's parents are interested in investing in the stock market to support his new finance career.  They are passionate about Green Energy and want to invest in DAQO New Energy Corporation.  DAQO (ticker DQ) makes silicone wafers for solar panels.  Steve's parents did no research on the company's performance but since they met at Dairy Queen they decided the coincidence was good enough to invest.  Before investing, Steve wants to analyze Green Energy stocks to see if DAQO appears to be a good investment.  
### Purpose:
In order to determine if DAQO is a good Green Energy stock to invest in Steve pulled 11 other Green Energy stocks to review performance.  He span stock performance over 2 years - 2017 and 2018.  He is interested in determining the stocks increase or decrease over the year as well as the number of shares traded throughout the day.  He will review the following metrics for all 12 Green Energy Stocks.

1. Daily Volume = total number of shares traded throughout the day &#8594; measures how actively a stock is traded
2. Yearly Return = percentage difference in price from the beginning of the year to the end of the year &#8594; measures if the stock price went up or down over the course of the year

Reviewing this data should help highlight which Green Energy stocks are performing poorly and which are performing well. 

## Results: 
### Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
After running the analysis DQ does not appear to be the best option for investment.  It had a strong performance in 2017 increasing almost 200% with low overall trading volume. 11 of the 12 Green Energy Stocks in 2017 all had positive returns in 2017.  However, DQ decreased 63% in 2018 with triple the trading volume.  Only 2 stocks returned positive investments for the 2 consecutive years analyzed.  These stocks were ENPH and RUN.  ENPH had 2 very strong years of performance returning 130% in 2017 and 82% in 2018.  RUN returned 6% in 2017 and 84% in 2018.  Steve should suggest to his parents to first consider ENPH and then RUN over DQ.  

Original Code 2017            |  Refactored Code 2017
:-------------------------:|:-------------------------:
![](https://github.com/lauras521/stock-analysis/blob/698fc1ed79afa3468d76f15d6067abb3b8b28d3c/Resources/VBA_Challenge_2017_Not_Refractored.PNG)  |  ![](https://github.com/lauras521/stock-analysis/blob/225a6692aa0f4ae5d9132c2828bee8dae3b0502a/Resources/VBA_Challenge_2017.PNG)


Original Code 2018            |  Refactored Code 2018
:-------------------------:|:-------------------------:
![](https://github.com/lauras521/stock-analysis/blob/698fc1ed79afa3468d76f15d6067abb3b8b28d3c/Resources/VBA_Challenge_2018_Not_Refractored.PNG)  |  ![](https://github.com/lauras521/stock-analysis/blob/225a6692aa0f4ae5d9132c2828bee8dae3b0502a/Resources/VBA_Challenge_2018.PNG)


Steve first wrote an Original Code and then updated a Refactored Code to try and speed up the analysis.  In the images above, you can see the Refactored Code ran faster than the Original Code.  

First all 12 tickers to be analyzed were identified in both codes.  Image below.  The main differences between the codes is highlighted in the table below.  If the code is too small to read the image can be clicked on to be enlarged. Steve updated his for loop format to iterate more efficiently in the Refactored Code to cut time out of the analysis. In the original code all 12 tickers were looped through and within each ticker loop all rows were looped through once to find tickerVolume, again to find startingPrice, and again to find EndingPrice.  Then data was outputed before moving to the next ticker.  In the new Refactored code the for loop isn't gone through as many times since the loop can identify when tickers have changed based on initializing the tickerIndex and not needing the AND statement to determine when ticker changes and to output the starting and ending price for a ticker.


<p align="center">
  <img src = https://github.com/lauras521/stock-analysis/blob/c18fcd628b998817db367378643ab54d5242e156/Resources/Initialize_Tickers.PNG>
</p>

 Original Code 2018            |  Refactored Code 2018
:-------------------------:|:-------------------------:
![](https://github.com/lauras521/stock-analysis/blob/c18fcd628b998817db367378643ab54d5242e156/Resources/Original_Code_For_Loop_and_If_Statements.PNG)  |  ![](https://github.com/lauras521/stock-analysis/blob/d13f2a968c731db79ca7deb551eef85fd3141f0d/Resources/Refractored_Code_For_Loop_and_If_Statements.PNG)

  
## Summary: 
### In a summary statement, address the following questions.
#### What are the advantages or disadvantages of refactoring code?
&emsp; ***Advantages of Refactoring Code***
* code can run faster
* code looks cleaner for analysis by others and sharing/editing purposes

&emsp; ***Disadvantages of Refactoring Code***
* it takes time to refactor code so if there isn't a big benefit for doing the work it could take up unnecessary time
* time is money in the workplace and lost time has real cost

#### How do these pros and cons apply to refactoring the original VBA script?
&emsp; ***Advantage of Refactoring the Stock_Analysis VBA Code***
* The code did run faster and if you were to want to analyze all tickers in the stock market or multiple years of data this time savings could pay off.

&emsp; ***Disadvantage of Refactoring the Stock_Analysis VBA Code***
* The code only took 1 second to run in the first place so while it did run faster with the Refactored code for this analysis the change wasn't necessary for the time tradeoff.

