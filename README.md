# Green Stock Analysis using excel VBA 

## Project Overview  

Steve has asked to perform of an analysis of green stocks to help his parents determine if it is worth investing in. To do the analysis, excels Visual Basic Applications (VBA) was used to determine the stocks daily volume and annual return for the years of 2017 and 2018.

## Purpose of Project

The purpose of the project is to:
A) Analyze multiple stocks at once 
B) Find a way to make the macro run as efficiently as possible

## Results  

The original macro has a nest loop with the outer loop  used to initialize the variable “totalVolume” to 0 while the inner loop performs the logic to find the total Volume for each of the stocks and the total return of each stock (expressed as a percentage).  The original macro ran in a time of 0.55 seconds for 2017 and 0.52 seconds for 2018 which results in an average time of 0.54 seconds. To make this code run faster instead of nested loops 2 separate loops were created. The first loop initialized the variable “tickerVolume” to 0,  the second loop performed the logic to find the total Volume and total return for each stock and the last loop tabulated the outputs the logic performed in the second loop. Prior to the execution of each loop, 4 new variables called “tickerIndex”, “tickerVolume” (variable type is an array), “tickerStartingPrice” (variable type is an array) and “tickerEndingPrice” (variable type is an array)  were created so when the loops executed the variables would hold the values the logic performs. Each of the arrays that were created were indexed to the tickers array which holds all the stocks needed for analysis through the variable called “tickerIndex”. Once the code was executed there was a drastic increase in performance as the coded executed in 0.059 seconds for 2017 and 0.063 seconds for 2018 which results in an average execution time of 0.061 seconds.

After calculating the total daily volumes and the returns of each green stock in 2017 and 2018 the conclusion that can be made is that Steve’s parents should invest in the stocks ENPH and Run. The RUN stock did exceptionally well as its return went up from 5.5% in 2017 to 84% in 2018 while the ENPH stock did do as well as its returns fell from 129.5% to 81.9% but it still nets a positive return compared to some of the other stocks that fell from 2017 to 2018 and netted a negative return 


## Summary 
- What are the advantages or disadvantages of refactoring code

    Code refactoring is a way to optimize existing code so it could improve performance and to restructure code, so it is more readable to fellow colleagues. The disadvantage to refactoring code is the potential of making the working code unusable due to error. One way to overcome this disadvantage would be to clone the working code and make the changes on 1 of the versions thus always have 1 working code.

  

- How do these pros and cons apply to refactoring orginal script 

    The advantage to refactoring the original code is that a lot of the logic is the same, so it became a matter of just reorganizing the loops and re-indexing the variables. The disadvantage though is that since there are a lot more variables to be kept track off potential for errors such as variable declarations increase.


## Run-Time Screen Shoots
- Orginal Code 
- Refactored Code 