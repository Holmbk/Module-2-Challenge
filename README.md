# Refactor VBA Code and Measure Performance
## Overview of Project
  The purpose of this Challenge  was an excellent gateway to programing in VBA and using Excel to desplay the data. Even though we did simular things as what could be done in Excel, it was a good way to see what we were programing could be easier and faster than just using Excel. This alsohelped make the programming easier to visualize and better to understand.
The background we were challenged to refactor Module2_VBA_Script to that it would loop through the stock ticker data one time versus doing eash ticker individually. WE had to create a variable with single or long data types. Write "for" loops and "if-then" statements.  We used logical andcomparison operators, an index to access data in an array, used nested loops, and reused code.  We had to debug our code and place comments
with in the code. We had to use visual and numeric formatting along with conditional formatting in order to show the rusults on a spreadsheet.We also measured the run time of the code and compaired it to the run time of the code only doing one stock ticker at a time.
 ## Results
 ## Results of Outcomes Based on 2017
   The Static Formatting created the table for all tickers data of 2017.  The Conditional Formatting highlighted all ticker's return that was 
   positive with a green cell.  All ticker's that had a negative return had a red cell, and all tickers that didn't have a return (value of 0)
   was left clear. All but one ticker had a positive return in 2017.  The one ticker had a negative value in 2017. So all but the one ticker
   appeared to be posible options for investment.
 ## Results of Outcomes Based on 2018
   The Static Formatting created the table for all tickers data of 2018.  The Conditional Formatting highlighted all ticker's return that was 
   positive with a green cell.  All ticker's that had a negative return had a red cell, and all tickers that didn't have a return (value of 0)
   was left clear. Only tow tickers had a positive return in 2018.  All other tickers had a negative value in 2018. So only the two tickers
   appeared to be posible options for investment. 
 ## Over All Results
  Both years (2017 and 2018) evaluations ran for the same amount of time of 0.0546875 seconds.  Running the evaluation for only one ticker versus
    12 tickers was the time of 0.2578125 seconds.  The one ticker evaluation ran so much longer than doing all tickers for the year at the same 
    time. This goes to show that the all tickers code was much more efficient. For investing, this analysis didn't cover enough years of data for
    a person to make a good judgement on their decision to invest in a company.  However just using our data only two tickers appear worthy for 
    investment. These tickers are "ENPH and RUN".  ENPH had the beter over all return with 129.5% in 2017 and 81.9% in 2018. RUN had a simular 
    2018 return of 84.0% but a lower return in 2017 of 5.5%
 ## Summary
  The advantages of refactoring code is time savings in writing new code. You only have to make adjustments to the older code to masauge it into creating your desired results.  The disadvantage of refactoring code is that the older code could be for a different version of software you are using and not operate correctly when run.  Also there may be errors left in it by the previous programmer that may blow up when you try to run the refactured code. The advantages of our refactored VBA script ran faster than the original and gave you the values for all the tickers.  A disadvantage of the original VBA script was its run time was longer for just one ticker as comtaired to the refactored VBA script.  It would take so much longer to create the same data for all of the tickers as the refactored VBA script.
 
 
