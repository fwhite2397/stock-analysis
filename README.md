# stock-analysis with VBA

## Overview of Project
This spreadsheet provides analysis of 12 stockes based upon their collected data, for calendar years 2017 and 2018.
The macros allow the user to produce summarized results for the either of the two year, but allowing the user
to input the year they which to report on. 

As part of this project, we first wrote code to calculate the results for the user and save the details in worksheet called,
"All Stocks Analysis". Although, the first implementation provide the expected results, we felt that the code could be optimize to
execute and generate the results in more timely fashion. The code submitted is a refactoring of the code and with the improved code 
that performs faster than the original.

## Results:
### Original Code Execution
The original code was written to write the results for each stock out to the "All Stocks Analysis" worksheet, after completing the loop for each stock. 
<img width="961" alt="2021-12-25 (4)" src="https://user-images.githubusercontent.com/95976771/147399445-660cbc83-2493-4842-8fce-8ac21845157a.png">

Here are the execution times for both 2017 and 2018:
<img width="230" alt="2021-12-25 (5)" src="https://user-images.githubusercontent.com/95976771/147399492-ab1df254-8473-4756-acc2-306e7142fa8a.png">

<img width="230" alt="2021-12-25 (6)" src="https://user-images.githubusercontent.com/95976771/147399518-b58e8f36-83b3-4cab-8386-5b3ed4084b27.png">

I'll prov

### New Code Execution
The new code was written to save the results for each stock in an array and after we have looped through ALL the stocks, write the result to reuslts to the "All Stocks Analysis" worksheet. 
<img width="961" alt="2021-12-25 (8)" src="https://user-images.githubusercontent.com/95976771/147399639-a0142d9a-2f12-4687-9449-d431735f0f9b.png">

This proved to be more efficient than writing out the results out one at a time. This is reflected in the execution times of the updated code. 
We verifed that the code produced the same results as the original (very important!) and executed faster.

Here are the execution times for both 2017 and 2018:
<img width="221" alt="2021-12-25" src="https://user-images.githubusercontent.com/95976771/147399594-0a2e5297-8836-480f-90a4-c7e6299dc918.png">

<img width="221" alt="2021-12-25 (2)" src="https://user-images.githubusercontent.com/95976771/147399622-9f3a793d-6adb-4b7b-8e8d-6df75ec89433.png">


## Summary:
### Advantages or disadvantages of refactoring code
The advantages of refactoring code is that you get to use new techniques to improve the experience of the customer. These newer technique may either
be unknown to the developer at the time when the original code was written, or not even available. But learning new ways of writing code that are more efficient 
can improve the performance of previously slow running processes. However, not all techniques will promise postive results, in the process of refacting, you 
could introduce either a bug or cause the process to run slower. 

### How do pros and cons apply to refacting the original VBA script
You must ensure that the effort put forth to refactor your code provides the same (or better) results than the original and must have a postive impact on the dependent user/process. In our example, refactoring the code provided both the same (expected) result and executed faster. This provided the outcome needed to just modifying the original version of code.
