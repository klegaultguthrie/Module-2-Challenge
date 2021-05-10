# MODULE 2 CHALLENGE - Refactor VBA Code and Measure Performance

## Project Overview
This project analyzes stock data from 12 stock tickers; specifically, the analysis compiles the total daily volume and the return generated from each stock over the years 2017 and 2018.


### Purpose
The purpose of this project is to refactor the original code with the intent of increasing the speed (efficiency) by which the VBA script is executed. Secondly, the performance of the stocks will be reviewed to ascertain if any conclusions can be drawn.


## Results
The original VBA scrips times are shown below; the scripts for 2017 and 2018 ran in .617 and .613 seconds respectively.

#### Original Run Times

![ORIGINAL All Stocks 2017](https://github.com/klegaultguthrie/Module-2-Challenge/blob/main/ORIGINAL%20All%20Stocks%202017.png)
![ORIGINAL All Stocks 2018](https://github.com/klegaultguthrie/Module-2-Challenge/blob/main/ORIGINAL%20All%20Stocks%202018.png)

Once the original script was refactored, the run times were successfully reduced to the times indicated in the images below. This was achieved via the creation of arrays for the ticker volumes, as well as the ticker start and end prices. The original scrip used if statements in lieu of arrays resulting in a longer execution time.


#### Refactored Run Times
![All Stocks 2017](https://github.com/klegaultguthrie/Module-2-Challenge/blob/main/All%20Stocks%202017.png)
![All Stocks 2018](https://github.com/klegaultguthrie/Module-2-Challenge/blob/main/All%20Stocks%202018.png)

Stock performance in 2017 was strong; all stocks had a positive return, with the exception of TERP which saw a -7.2% return. Comparatively, 2018 stock performance was poor, as all stocks had negative returns with the exception of ENPH and RUN, which posted healthy returns (greater than 80%). Let's assume the data set for both 2017 and 2018 is representative of the US stock market. Given the largely favorable returns generated in 2017, we can infer the economy on a whole is growing and that investors are confident in the markets resulting in increased investment, stock growth and increased stock returns.
Subsequent to 2017, we can infer that the economy was impacted negatively by one or more events. We can also infer that the two stocks that reported gains were some how not impacted by economic factors to the same extent as the rest of the market. Realistically however, analyzing the stock market would include inumerably more tickers than 12 (as was the case in this scenario). Our analysis is limted by the fact that we do not know what industry these stocks belong to, nor how long they have been actively traded for. 


## Summary

The original VBA code was easier for me as a new coder to understand and follow, given my familiarity with the 'If Statement' function in Excel. The main disadvantage is that the code takes longer to execute.

The refactored code was more difficult for a beginner to interpret; however, the code was executed in less time. The latter would be of increasing significance if the number of tickers analyzed were to increase significantly.

The benefits of refacting code are:
1) This is a effective way to become more familiar with reading and interpreting code from other sources
2) The ability to rewite code in order to increase its efficiency is desirable in the context that computing power utilizes ressources.
3) Peer review in a common part of code refactoring and a common ask of the industry.


Some of the drawbacks of refactoring code:
1) Rewriting code that someone else created can be time consuming and challenging.
2) The complexity and cost of refactoring code may be greater than writing the code from scratch.
3) Refactoring code may introduce bugs, and there may not be enough time if the delivery schedule is tight.
https://www.c-sharpcorner.com/article/pros-and-cons-of-code-refactoring/

