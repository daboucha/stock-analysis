# Stock Analysis

## Overview of Project

Steve has recently graduated with his finance degree, and is looking to advise his parents on their investments. Steve’s parents are interested in investing in green energy stocks, and currently have all of their money invested in DAQO New Energy Corp (DQ). Steve has promised to look into DAQO stock for his parents, but he is concerned about diversifying their funds. Steve wants to analyze some more green energy stocks in addition to DAQO stock, he has created an Excel file containing the data he would like to analyze, and has asked for my help to conduct the analysis.  I will be using Visual Basic for Applications, or VBA, to automate tasks and perform our analysis on all of the available stock options.

## Resources
- Data Source: green_stocks.xlsx
- Software: Microsoft Excel 2019

## Tasks Accomplished in this Module:

- Create a VBA macro that can trigger pop-ups and inputs, read and change cell values, and format cells
- Use for loops and conditionals to direct logic flow
- Use nested for loops
- Apply coding skills such as syntax recollection, pattern recognition, problem decomposition, and debugging

## Analysis Results

After refactoring the code of the original stock analysis workbook, I was able to reduce looping redundancies and reduce the program’s run time. Below is a picture of the captured run time of program before and after the refactoring for the Stock Analysis of 2017.

![2017_Run_Time](assets/VBA_Challenge_2017.png), ![2017_Refactored_Run_Time](assets/VBA_Challenge_Refactored_2017.png)

As you can see, the refactoring has done its job of reducing the program’s run time while having no impact on the output. You will notice a similar decrease in run time when using the refactored program to analyze the data from 2018 as well. 

![2018_Run_Time](assets/VBA_Challenge_2018.png), ![2018_Refactored_Run_Time](assets/VBA_Challenge_Refactored_2018.png)

The decrease in the run time is only marginal in the analysis of 12 stocks. But the refactored program is able to perform the same analysis on much larger data sets with any number of stocks, unlike the original version which would require additional code to be entered manually for every new stock ticker in the data set.

## Analysis Summary

1.	The most obvious advantage of refactoring is that it leads to better quality code. It may help the program run faster, but more importantly it improves the design and makes the software easier to understand. Refactoring will also make the program easier to debug. And by storing all arbitrary numbers into variables it becomes much easier to edit the code due to changes in the data set format. The disadvantages to refactoring code is that it is time spent on a program that is already functional, and you may unintentionally cause bugs while refactoring which turns your previously working program into one that doesn’t work.
2.	In the case of our Stock Analysis program, the biggest advantage is that we can increase the data set to include additional stocks and the program can produce the analyzed output with no changes required to the code. Also, if the code required editing due to changes to the data set’s format, such as the stock’s ticker, daily volume, or closing price being moved to different locations, then we would save on time since the number of changes required on the refactored code are much less than on the original code. The only disadvantage to refactoring this program was the time spent on editing the code, but that was a small price to pay considering the improvements.