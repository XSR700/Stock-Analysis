# Stock-Analysis

##Overview of Project
This Project is created to help Steve choose a Renewable Energy stock for his parents. His parents want to invest all of thier money in to DAQO New Energy Corp (DQ). However Steve's parent have not done any prior research. This project will compare a few other green energy stocks with DQ and output the total volume of each stock and thier performance. This information will be represented using Excel's Visual Basic Appliation (VBA). There is a written Macro code that will generate the informationwith a clock of a button.

This Project cotains an original script Macro and refactored script Macro. The comparisson between the two will also be presented and discussed for comparrison.  

##Results
Each script was written with a timer that recorded the total run time. After running the originalscript for year 2017 and 2018, we see that the total run time was the same. Which makes sense because each year had the same amount of rows that it looped over.

![Original_Script_Run_Time_For_2017](https://github.com/XSR700/Stock-Analysis/blob/main/Resources/Original%202017%20Time.PNG)

![Original_Script_Run_Time_For_2018](https://github.com/XSR700/Stock-Analysis/blob/main/Resources/Original%202018%20Time.PNG)

Now we look will look into the performance of the refactored scripts for years 2017 and 2018. We see that the run time for 2018 was a bit faster.

![Refactored_Script_Run_Time_For_2017](https://github.com/XSR700/Stock-Analysis/blob/main/Resources/Refactored%202017%20Time.PNG)

![refactored_Script_Run_Time_For_2018](https://github.com/XSR700/Stock-Analysis/blob/main/Resources/Refactored%202018%20Time.PNG)

Regardless of the fact that 2018 was a few miliseconds quicker, we can say that they run about the same speed. The order of the run may effect the speed. For example,year 2018 was run after running 2017 here, but vice versa the year 2017 was a few miliseconds faster instead. Therefore the order of the run may effect the code speed marginally.

In any case, the most important factor here is the set up of the code. We see that the refactored script is faster by about  
