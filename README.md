# stock-analysis
UofT Data Analytics Boot Camp Module 2 - VBA
## Overview of Project
* The raw data contains energy stock information such as price and volume which will be used to analyze and compare the performance of these various stocks. We have created a VBA script to perform this analysis, and have re-factored the code to make it more efficient. The analysis returns the total daily trading volume along with percentage return (calculated by comparing the change in the stocks' ending price from the starting price for the specified year). This analysis is done for the years 2017 and 2018.
## Explain the purpose of this analysis.
* Our client, Steve, has asked us to perform an analysis on other energy stocks to compare their performance against the stock his parents want to invest in, ticker "DQ". Steve is interested to see the performance of other stocks in order to determine whether there are other stocks he can diversify into. 
## Results: 
### Compare the stock performance between 2017 and 2018
![VBA_Original 2017.PNG] https://github.com/BBBrian1124/stock-analysis/blob/Module-2_Challenge_UofT_Data_Analytics_Bootcamp/Resources/VBA_Lesson_2017.PNG
* The analysis returns the percentage change between the stock's ending price for the year and the starting price. Looking at the data for 2017, it would appear that it was a good year in the stock market with all stocks having a positive return in this industry. A look at DQ's leading 199% return may suggest that this is a great stock to invest in, however, such a high return may also suggest that we would be buying the stock at an overvalued or high entry point. This kind of large growth with stocks can also mean that the period of high growth has passed, and that we'd expect to see slower or more steady growth in the future as opposed to this exponential growth, meaning it may be unlikely that we see this same percentage of returns. There may have been other factors that caused such a large industry wide growth that need to be looked into as well. Ideally, we'd want to buy into the stock either before this type of high growth, or buy into a stock we can expect steady growth time over time so that we can maximize our returns.

![VBA_Original 2018.PNG] https://github.com/BBBrian1124/stock-analysis/blob/Module-2_Challenge_UofT_Data_Analytics_Bootcamp/Resources/VBA_Lesson_2018.PNG
* A look at the year after shows that we may have seen a stock market correction, with almost all stocks having a negative return in 2018. Again, DQ is leading here with the largest % loss at 63%. A market wide decrease like this could suggest trouble in this industry or a change in views in this industry. Again, other factors would need to be looked at to investigate what caused this change since it is industry wide (i.e. there may have been regulatory changes, substitutes, etc. that harmed this industry).

![Stock Analysis.PNG] https://github.com/BBBrian1124/stock-analysis/blob/Module-2_Challenge_UofT_Data_Analytics_Bootcamp/Resources/Stock%20Analysis.PNG
* What may be better to look at is the price across the years. A look at this shows that the only two tickers that have been constantly growing are: ENPH and RUN. Additionally, the change in volume should also be considered. Volume in the stock market has an impact on the volatility of the stock (i.e. how much the stock price fluctuates). Generally speaking, the lower the volume of the stock the more volatile that stock is, and the greater the volume the less volatile it is. From this visual, we can see that ENPH and RUN are not only showing steady increases over time in their stock price, but also increase in volume. This suggests that these two stocks may be "safer" as their prices appear to be growing steadily and are less likely to have price fluctuations. Further analysis/research is needed with regards to this industry, however, these two tickers may be worth a look in addition to (or instead of) "DQ".

### Compare the execution times of the original script and the refactored script.
![VBA_Refactored_2017.PNG] https://github.com/BBBrian1124/stock-analysis/blob/Module-2_Challenge_UofT_Data_Analytics_Bootcamp/Resources/VBA_Challenge_2017.PNG

![VBA_Refactored_2018.PNG] https://github.com/BBBrian1124/stock-analysis/blob/Module-2_Challenge_UofT_Data_Analytics_Bootcamp/Resources/VBA_Challenge_2018.PNG
* The visuals in the section above show the time it took the code to execute prior to being refactored. These two screenshots show the time it took the code to execute after being refactored. This is about an 85% decrease in time taken. 

![Original Code.PNG] https://github.com/BBBrian1124/stock-analysis/blob/Module-2_Challenge_UofT_Data_Analytics_Bootcamp/Resources/Original%20Code.PNG
* The reason has to do with the structure of the loops. This loop is used to assign the total volume, the starting price and the ending price to each of the 12 stock tickers. The loop starts at i = 0 then it executes a nested loop which looks through all the rows in the data sheets (starting at row 2 to the last row for a total of 3012 rows). The loop then starts again at i = 1, and executes the nest loop again. This essentially results in 33132 "executions" (11 * 3012).

![Refactored_Code.PNG] https://github.com/BBBrian1124/stock-analysis/blob/Module-2_Challenge_UofT_Data_Analytics_Bootcamp/Resources/Refactored%20Code.PNG
* In the refactored code, the structure of the loop changes. By using the variable tickerIndex, we are able to assign the total volume, the starting price and the ending price to the tickerIndex, which we are increasing within this loop as well (rather than having another loop). The loop now only runs through the rows in the data sheets, and with this structure, we are able to assign the values to the 12 stock tickers within one loop, rater than a nested loop. This results in 3012 "executions" for the same function as opposed to 33132 in the original code.

## Summary: In a summary statement, address the following questions.
### What are the advantages or disadvantages of refactoring code?
### Advantages:
* Efficient: we are able to re-use something that we know has worked in the past, therefore we don't need to research/brainstorm the solution
* Structure: we have a skeleton/structure for working code so we won't have to brainstorm the code flow from fresh and we have an outline to follow
* Improvement: we can look at what is already done and see if we can improve it, rather than thinking of an idea from fresh
### Disadvantages:
* Requires understanding: it requires us to understand what the prior code is doing, this can be difficult if the prior code is not documented well
* Requires same/similar scenario: using prior code requires the scenario to be similar/same, otherwise it may result in too many changes, which at that point it may be easier to write the code from fresh
* Assumes prior code works: needs to check for any errors or remove code that doesn't apply in our scenario
### How do these pros and cons apply to refactoring the original VBA script?
* Cons: In our scenario, we were provided the original code to refactor. In this scenario, the purpose of the original code was for the same purpose and was the code that we used in our module lesson, therefore the cons were lessened since the code applies to the same scenario and we have tested it to know it works. One of the cons that I had difficulties with was understanding the structure. The order of the instructions made it a bit difficult to conceptualize the flow of the code. The instructions advised to use a variable tickerIndex which doesn’t change values until the code in instruction 3d, therefore, I had a hard time to visualize how the values in each of the prior code would be filled while I was following along chronologically. I was able to overcome this by following the instructions backwards, starting from step 4 which better helped me conceptualize how the data needed to be presented and how the tickerIndex value changed. Additionally in troubleshooting my "infinite loop" challenge (see appendix for more details), I created a table that shows the flow of the loop/code in non-coding language and created a new sheet which presented this in the format of Excel formulas, both which I am much more familiar with. By seeing the flow of the code in these ways I was able to better understand what was happening with the code.
* Pros: Once I understood the structure better, I was able to experience the pros of refactoring as there was essentially only one major change (which was to the loop structure as explained in the "Results" section), and I was able to just slightly tweak the code from our module lesson to get the code to work. Again, we also saw an improvement in the time it took the code to run.

### Appendix:
[VBA Challenge File] https://github.com/BBBrian1124/stock-analysis/blob/Module-2_Challenge_UofT_Data_Analytics_Bootcamp/VBA_Challenge.xlsm
* Contains the raw macro enabled Excel file with the code and analysis, the VBA code for the challenge is under module 2 of the file, the code for the lesson is in module 1

[Challenges Faced] https://github.com/BBBrian1124/stock-analysis/blob/Module-2_Challenge_UofT_Data_Analytics_Bootcamp/Challenges%20Faced.docx 
* Contains more details about the challenges faced during this project

[Link to Repository] https://github.com/BBBrian1124/stock-analysis/tree/Module-2_Challenge_UofT_Data_Analytics_Bootcamp
