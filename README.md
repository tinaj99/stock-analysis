# VBA Challenge

## Overview of Project
We are analyzing multiple stocks in a workbook using Visual Basic scripts for two different years. It is required to refactor the visual basic code to make it run more efficiently with lower execution times.

### Purpose
This project involves refactoring Visual Basic scripts to achieve faster execution times; to scale up better when dealing with large amounts of data.

## Analysis

### Comparison of Stock Performance Between 2017 and 2018

The figures below illustrate the Stock Performance in the years 2017 and 2018, after refactoring the existing VBA code.

##### Stock performance in 2017
![](resources/Stock_Performance_2017.PNG)

##### Stock Performance in 2018
![](resources/Stock_Performance_2018.PNG)

We can see that overall returns were significantly better in 2017 as compared to 2018. In 2017, 11 stocks showed positive percentage returns with 1 showing negative returns; whereas in 2018 only 2 stocks had positive returns with 10 being in negative territory.

Only two stocks, ENPH and RUN had positive returns in both years.

Only two stocks, RUN and TERP, performed better in 2018 vs 2017; all other stocks had worse returns in 2018 vs 2017. RUN had positive returns in both years, the with 84.0% return in 2018 vs 5% in 2017. TERP had negative returns in both years; but the losses were less in 2018.

### Comparison of Execution times of original and refactored script

The figures below illustrate the execution times for 2017 and 2018, before and after refactoring the code.

####Execution Time Comparison for 2017 Before and After Refactoring

Execution time reduced by 66.4% after refactoring the code (~0.222 seconds vs ~0.664 seconds).

![](resources/Execution_Times_2017_Before_Refactoring.PNG)

![](resources/Execution_Times_2017_After_Refactoring.PNG)


####Execution Time Comparison for 2018 Before and After Refactoring

Execution time reduced by 67.0% after refactoring the code (~0.218 seconds vs ~0.664 seconds).

![](resources/Execution_Times_2018_Before_Refactoring.PNG)

![](resources/Execution_Times_2018_After_Refactoring.PNG)

### Analysis of execution time improvements

The main reason for the performance improvement and resulting reduced execution time is the removal of the additional loop as illustrated by the code snippet Before and After the refactoring.

#### Before Refactoring

![](resources/Two_Loops_Before_Refactoring.PNG)

#### After Refactoring

![](resources/Single_Loop_After_Refactoring.PNG)

## Summary

### What are the advantages and disadvantages of refactoring code?

Some of the advantages of refactoring code are that we can make the code perform better, potentially using lesser system resources such as memory and improving the execution times by making the logic more efficient from a compute perspective;for example by eliminating extra for loops.

The disadvantages of refactoring is that we need to re-test a lot of functionality, and we could introduce bugs if we're not able to completely test the functionality - for example if the CI/CD pipeline and automated tests are not comprehensive enough.


### How do these pros and cons apply to refactoring the original VBA code

We have reduced execution times by ~66%, which is a signficant improvement, by refactoring the code. At the same time, we had to spend time to understand existing logic which could have been written by a third person, potentially introducing bugs and spending time to debug the code. Secondly, we had to re-run the test and validate that there is no regression in the correctness of the results.


