# stock-analysis

## Project Overview
At the behest of Steve, I was tasked to expand the dataset given to me to include the entire stock market over the last few years by refactoring the VBA code to loop through all the data.  Then, I had to determine if doing so speeds up the execution of the VBA script.

## Resources
- Data Sources: green_stocks.xlsx, VBA_Challenge.xlsm
- Software: Microsoft Excel

## Results
(/Images/VBA_Challenge_2017.png)
For 2017, the stock performance was overall robust, with every stock except for TERP (-7.2%) yielding a positive return.

(/Images/VBA_Challenge_2017_clock.png)
The code for the 2017 stock data took 0.695 seconds to execute.

(/Images/VBA_Challenge_2018.png)
By contrast, for 2018, the stock performance was just as overall terrible, with only ENPH (+81.9%) and RUN (+84.0%) yielding positive returns.

(/Images/VBA_Challenge_2018_clock.png)
The code for the 2018 stock data took 0.703 seconds to execute.

## Summary
One of the advantages of refactoring code is being able to modify it to suit different desired tasks or to perform existing ones more efficiently, while one of the disadvantages of doing so is introducing additional complications that may cause bugs and errors that have to be troubleshot, or unknowingly degrading the performance of the code relative to its original form.  As it relates to the VBA script worked with here, the code has definitely worked better compared to before and has become more convenient to use.
