# Refactor VBA Code and Measure Performance
## Purpose
  This Challenge  was an excellent gateway to programing in VBA and using Excel to desplay the data. Even though we did simular things as 
  what could be done in Excel, it was a good way to see what we were programing could be easier and faster than just using Excel. This also
  helped make the programming easier to visualize and better to understand.
## Analysis
   WE were challenged to refactor Module2_VBA_Script to that it would loop through the stock ticker data one time versus doing eash ticker 
   individually. WE had to create a variable with single or long data types. Write "for" loops and "if-then" statements.  We used logical and
   comparison operators, an index to access data in an array, used nested loops, and reused code.  We had to debug our code and place comments
   with in the code. We had to use visual and numeric formatting along with conditional formatting in order to show the rusults on a spreadsheet.
   We also measured the run time of the code and compaired it to the run time of the code only doing one stock ticker at a time.
 ## Analysis of Outcomes Based on 2017
   The Static Formatting created the table for all tickers data of 2017.  The Conditional Formatting highlighted all ticker's return that was 
   positive with a green cell.  All ticker's that had a negative return had a red cell, and all tickers that didn't have a return (value of 0)
   was left clear. All but one ticker had a positive return in 2017.  The one ticker had a negative value in 2017. So all but the one ticker
   appeared to be posible options for investment.
 ## Analysis of Outcomes Based on 2018
   The Static Formatting created the table for all tickers data of 2018.  The Conditional Formatting highlighted all ticker's return that was 
   positive with a green cell.  All ticker's that had a negative return had a red cell, and all tickers that didn't have a return (value of 0)
   was left clear. Only two tickers had a positive return in 2018.  The other tickers had a negative values in 2018. So only two tickers
   appeared to be posible options for investment. 
 ## Challenges and Difficulties Encountered
   Challenges were creating the for loops and if-then statements.  Typo's and spelling gave challenges in getting things to run.  Syntacs were 
   not to bad since it was coverd in class well.
 ## Results
    Both years (2017 and 2018) evaluations ran for the same amount of time of 0.0546875 seconds.  Running the evaluation for only one ticker versus
    12 tickers was the time of 0.2578125 seconds.  The one ticker evaluation ran so much longer than doing all tickers for the year at the same 
    time. This goes to show that the all tickers code was much more efficient. For investing, this analysis didn't cover enough years of data for
    a person to make a good judgement on their decision to invest in a company.  However just using our data only two tickers appear worthy for 
    investment. These tickers are "ENPH and RUN".  ENPH had the beter over all return with 129.5% in 2017 and 81.9% in 2018. RUN had a simular 
    2018 return of 84.0% but a lower return in 2017 of 5.5%
 ##
 
 
