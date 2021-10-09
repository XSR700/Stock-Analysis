# Stock-Analysis

## Overview of Project

This Project is created to help Steve choose a Renewable Energy stock for his parents. His parents want to invest all of their money into DAQO New Energy Corp (DQ). However, Steve's parents have not done any prior research. This project will compare a few other green energy stocks with DQ and output the total volume of each stock and their performance. This information will be represented using Excel's Visual Basic Application (VBA). There is a written Macro code that will generate the information with a click of a button.

This Project contains an original script Macro and refactored script Macro. The comparison between the two will also be presented and discussed.  

## Results

The two tables below display all green energy stock performances that Steve considered for years 2017 and 2018 respectively. The third column represents the percentage of return from the starting price to the ending price. We notice right away that in 2017, almost all of the stocks had a profitable yield while in 2018 most of the stocks lost value. It seems something happened in 2018 that drove the renewable energy industry down in value.

![2017 Stock Performance](https://github.com/XSR700/Stock-Analysis/blob/main/Resources/ALL%20STOCKS%202017.PNG)

![2018 Stock Performance](https://github.com/XSR700/Stock-Analysis/blob/main/Resources/ALL%20STOCKS%202018.PNG)

The stock "DQ" that Steve's parents considered buying grew in value in 2017. In fact, "DQ" showed the biggest return of 199% in 2017 among the others. At the end of 2018, "DQ" fell in value by 62% which gives it a total growth of 137% over two years. However stock tickers "RUN" and "ENPH" grew in value throughout 2017 and 2018. the ticker "RUN" grew 5.5% in 2017 and 84% in 2018 giving a total of 89.5% over two years. the ticker "ENPH" grew 129.5% in 2017 and 82% in 2018 which gives a total of 211.5% over two years. We can conclude that based on our data and analysis that stocks "RUN" and "ENPH" are probably safer investments than "DQ". "ENPH" should hold the majority of the investment because it stayed profitable throughout 2018 and almost doubled in value over two years. 

Each script was written with a timer that recorded the total run time. After running the original script for the years 2017 and 2018, we see that the total run time was approximately the same with a difference of roughly 0.01 seconds. This should be expected since the number of rows for the years 2017 and 2018 are the same. 

![Original_Script_Run_Time_For_2017](https://github.com/XSR700/Stock-Analysis/blob/main/Resources/Original%202017%20Time.PNG)

![Original_Script_Run_Time_For_2018](https://github.com/XSR700/Stock-Analysis/blob/main/Resources/Original%202018%20Time.PNG)

Now we will look into the performance of the refactored scripts for years 2017 and 2018 below. We see that the run time of refactored scripts from both 2017 and 2018 was faster than the original scripts. The refactored script is roughly 0.05 seconds faster. 

![Refactored_Script_Run_Time_For_2017](https://github.com/XSR700/Stock-Analysis/blob/main/Resources/Refactored%202017%20Time.PNG)

![refactored_Script_Run_Time_For_2018](https://github.com/XSR700/Stock-Analysis/blob/main/Resources/Refactored%202018%20Time.PNG)

Even after completing several more runs and comparing the times, the timer shows the refactored script to be faster than the original. We can conclude that the refactored script runs faster.  

# Summary

## What are the advantages or disadvantages of refactoring code?

Refactoring a code is a common coding process. This allows the existing code to run more efficiently and faster. The main advantage of refactoring a code is run time. For programs that run through millions of loops and iteration, code run time is very important for the end-user. A refactored code runs fewer iterations to give results faster.

A disadvantage of refactoring a code is that, similar to the original script, it still holds the issue when running for the first time. When running a macro the first time, it will run slower than the next time. This is true to a refactored code still. This happens because the computer has to allocate all the resources and store the new variables. Ideally, a code should run as fast as possible in every run, and thus this is a drawback.   

## How do these pros and cons apply to refactoring the original VBA script?

When the refactored code ran the first time, the timer showed approximately the same time as the original script after multiple runs. This applies to both years. The screenshots above are results from non-first-time runs. The refactored code itself runs faster than the original, but both methods run slower on their first runs. 

