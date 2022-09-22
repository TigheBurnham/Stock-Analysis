# VBA Stock-Analysis

Visual Basic for Applications is an internal programming language designed to allow users to create and run macros, so they can achieve options that are not normally available in MS Office applications. This project delves into the power of macros and automation especially when analyzing large datasets like stock prices and volumes over a period of time.

## Overview of Project

This project has two primary goals: the first objective is to produce a dataset to help a client determine if investing in a certain dozen stocks between 2017 and 2018 is a worthwhile investment. The second goal of the project was to refactor the original code given to the client to see if the program would run more efficiently, and then communicate the results. 

### Purpose of Refactoring 

Refactoring is the process of restructuring code so it will internally run more efficiently without changing any of the external results. Writing more efficient code is important because it becomes easier to understand, troubleshoot, and will decrease run times which can be exponentially more important with very large datasets. 

### Data 

There are 12 different stocks, the original excel spreadsheets are organized in a way to show the stock's ticker, opening price, closing price, and volume. 

### How the code was made more efficient

The first goal was to figure out how to write the code to produce more efficient results. In the original code there was a nested for loop otherwise a loop in a loop. For loops are notorious for being the limiting behavior of a program, and can drastically increase Big O run times. Therefore, to make the code run more effectively we removed the nested for loop and utilized arrays to store the data of the starting and ending stock prices. 

#### Nested Loop Code

![Nested Loop](https://user-images.githubusercontent.com/112028534/191863963-570bbb9f-3420-417e-bb96-05b6bba55c31.PNG)

#### Code with Arrays
![Loop with arrays](https://user-images.githubusercontent.com/112028534/191864099-44ca175b-97ba-4803-a9c0-b89ded744d4e.PNG)

## Runtime Results

The two primary concerns are to observe if the refactored code has a shorter run time than the original, and to see if the new code actually produces the same results as the original.

### Original Code Runtime with Nested For Loop

![Unrefactored code time](https://user-images.githubusercontent.com/112028534/191857229-3c63dba7-f464-4021-9d76-24a8ee7ef13e.PNG)

### Refactored Code Runtime with Array

![Refactored Code Runtime](https://user-images.githubusercontent.com/112028534/191855753-adbec142-6e58-4337-ae0d-258a6f5764ee.PNG)

## External Results

### Original Results

![VBA_Challenge_2018_Original](https://user-images.githubusercontent.com/112028534/191857214-a0f7eca4-aace-4495-b69a-af51e372cc5b.PNG)

### Refactored Results

![VBA_Challenge_2018](https://user-images.githubusercontent.com/112028534/191855105-31687347-142b-44e5-a115-ecb22c1091d5.png)

As we can see the refactored code produces the same results in a shorter runtime; therefore, it appears the original code was successfully refactored.

## Summary

###  Advantages and Disadvantages of Refactoring Code

Again the major advantage of refactoring is the decrease in macro run time. In this analysis, the new code ran  faster than the original. A disadvantage of refactoring code is that after a program has been deployed and there have been updates and extensions added to the original, refactoring may break those updates or extensions. Therefore, it is easier to refactor code right after it has been deployed before any updates or extensions have been added.

### Pros and Cons of Refactoring the Original VBA Script

Refactoring takes extra time. Was this extra time worth it? Instead of refactoring the code more time could have been spent finding other stocks that could have had a greater return on the investment than the stocks chosen by the client. It appears to be the client's primary goal is to find a group of stocks that meet his financial goals, not to have the dataset macro to run more efficiently. 
