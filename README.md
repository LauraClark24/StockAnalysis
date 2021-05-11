# StockAnalysis
## Overview of Project
### Purpose

Steve's parents have hired him to help them invest in green energy. We are helping Steve analyze their current investment in DQ as well as other companies they may be interested in. Instead of just doing the analysis for Steve, we designed for him a VBA macro, a piece of programming in Excel, to analyze data from the current sheets as well as any more he adds in the future.

## Analysis of Outcomes
### Results

After spending multiple days fine tuning our macro, these reesults were acquired.  

![VBA_Challenge_2017](https://user-images.githubusercontent.com/83182353/117887825-e737a200-b276-11eb-93a0-09c02b2f1d69.png)

Of our current sheets, 2017 is the first to look at.  A brief glance will show us that almost all of our compannies for assessments did well in that year.  A closer look will reveal that the highest returns was DQ that year. 

![VBA_Chalenge_2018](https://user-images.githubusercontent.com/83182353/117887901-ff0f2600-b276-11eb-82c7-3ae7b94caf6d.png)

The next year tells a very different story with almost all of the stocks dropping. Only two companies, ENPH and RUN, managed to have an increase. One of the more discouraging facts to be seen is that the largest decrease was in our previously succcessfull DQ, with a 62% loss.

### Conclusion

The best suggestion I could make is for Stevee to get another year or two after 2018 to analyze.  The stockmarket can fluctuate greatly and only two data points does not give a full picture. If more data cannot be aacquired to give financial advice, the conclusion from the data would be to withdraw from DQ. The fluctuations between the two years indicates a lack of stability that is not recommended. ENPH had the largest return acrosss both years, but RUN had an increase in return from 2017 to 2018. Either of these two compaanies seems like better investments than DQ from this analysis.

## Summary
### Refactoring

The entire purpose of refactoring code is to rethink, reorganize, and optimize whatever you are working on. This can be extraordinarily difficult for people who cannot come up with creative solutions or those who cannot see when their idea is not better than the current code.  Anohter thing that can make the process more challenging is when you are dealing withh code that someone else has written. Having to figure out why things are where they are or whether if a specific line of code is necessary is a process later editors face with only comments in the code to guide them.  Generally, it is usful to go back through code to remove unnecessary lines or make loops more efficient.  There are constant challenges, but the benefits can be substancial and often make it worthwhile. 

### Stock Analysis

Our original VBA script was relatively inefficient, using nested loops to go through our entire 3,000 data points twelve separate times to get all the information.  The simplest optimization question from there is 'can we do all that in one go?' Looking at our final code, the answer to that is yes.  In order to make our solution for that question work properly, our data sets have to fit a very specific format that our code is made for. Steve should be aware that the refactored code requires that those exact twelve stocks to be in the given order or their data will not be read and all subsequent data will also fail to be acquired. Given that this refactoring increased the speed of thee code to merely 1/5 of the time, that seems like a reasonable tradeoff.
![firstTimedRun18](https://user-images.githubusercontent.com/83182353/117887992-249c2f80-b277-11eb-9232-761c136f3523.png)

![VBA_Chalenge_2018](https://user-images.githubusercontent.com/83182353/117888005-28c84d00-b277-11eb-9930-42ebfdcc2b3d.png)

