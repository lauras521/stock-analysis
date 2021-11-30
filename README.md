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

Original Code 2017            |  Refractored Code 2017
:-------------------------:|:-------------------------:
![](https://github.com/lauras521/stock-analysis/blob/698fc1ed79afa3468d76f15d6067abb3b8b28d3c/Resources/VBA_Challenge_2017_Not_Refractored.PNG)  |  ![](https://github.com/lauras521/stock-analysis/blob/225a6692aa0f4ae5d9132c2828bee8dae3b0502a/Resources/VBA_Challenge_2017.PNG)


Original Code 2018            |  Refractored Code 2018
:-------------------------:|:-------------------------:
![](https://github.com/lauras521/stock-analysis/blob/698fc1ed79afa3468d76f15d6067abb3b8b28d3c/Resources/VBA_Challenge_2018_Not_Refractored.PNG)  |  ![](https://github.com/lauras521/stock-analysis/blob/225a6692aa0f4ae5d9132c2828bee8dae3b0502a/Resources/VBA_Challenge_2018.PNG)


Steve first wrote one set of code ("Original Code") and then updated a Reractored Code to try and speed up the analysis.  You can see the main differences between the codes below.  He updated his for loop format to iterate more effiicnetly in the Refractored Code to cut time out of the analysis.  In the images above you can see the Refractored Code ran faster than the Original Code.  



## Summary: 
### In a summary statement, address the following questions.
#### What are the advantages or disadvantages of refactoring code?
&emsp; Advantages of Refractoring Code
* faster

&emsp; Disadvantages of Refractoring Code
* hard to edit and keep track of

#### How do these pros and cons apply to refactoring the original VBA script?
