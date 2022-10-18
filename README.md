# VBA Challenge

## Overview of Project

### Purpose

We are asked to refactor VBA code and measure performance of our macros against stock data during 2017 and 2018. 


## Analysis

We started with a basic VBA script to run our analysis. This gave us information about the stock DQ. When determined that this stock was not performing well, we looked to run analysis against all the stocks from the data. We included custom formatting, header names, loops and nested loops to run across the entire spreadsheet in worksheet 2018. we also included a timer function to test how long our macro took to run once we provided input.
link to the data file: [filename](/VBA_Challenge.zip)

The first image was taken of the code run before refactoring. As you can see it took approximately 1.8 seconds to run data for 2018.

-  ![This is an image](/resources/2018AllStocksAnalysis.png)

### Analysis after Refactoring

A VBS script file was provided and we needed to use it to refactor our VBA code. The goal of refactoring is to improve the performance of the code by eliminating reducing the number of commands in the code. Be reducing and cleaning up the code, the code should run faster without any errors.

After refactoring I ran the macro and as you can see it only tool .32 seconds to run for 2018 data. 
    -  ![This is an image](resources/Refactor2018AllStocksAnalysis.png)

### Analysis of Stocks



### Challenges and Difficulties Encountered

I had to be careful with some of the code. I spent quite a lot of time trying to avoid misplacing the 'End If' and 'Next i' for the 'For' loops. 

## Results

