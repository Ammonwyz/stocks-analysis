# Stocks-Analysis
  Performing the analysis on All stocks for Steve's family by using VBA. 

## Overview of Project:
  The purpose of this project is that to extract and analyze the data of stocks in 2017 and 2018 through VBA. The results can us to understand the implementation and expected development of stocks. Help Steve's parents know how to make the right investments.
The original VBA macro was only for the target processing data of 12 stocks. For efficient analysis of more data, it is necessary to refactor.

## Results:
  By refactoring the code, it can be seen that the execution time has decreased from 0.7 and 0.8 seconds to 0.2 and 0.3 seconds. In the original macro, the code repeatedly read the data to analyze and output it as instructed. In the improved program, because of the index used, only one side of the data needs to be read, and then different analyses and outputs are performed, so the running time is saved. The new macro is good for running more stocks, such as hundreds or thousands.

* Original script

2017

![image_name](https://github.com/Ammonwyz/stocks-analysis/blob/15bae4f73eb008fb55e5e5c15b56be060519ccd0/Original%20VBA%202017.png)

2018

![image_name](https://github.com/Ammonwyz/stocks-analysis/blob/0a46e854ab4fb130cc9fb1e86de0ceced8208daf/Original%20VBA%202018.png)


* Refactored script

2017
![image_name](VBA_Challenge_2017.png)

2018
![image_name](VBA_Challenge_2018.png)



## Summary
1. Advantages and disadvantages of refactoring code in general

*Advantages:*
It runs fast and can analyze big data.

*Disadvantages:*
Code requires high memory and is not easy to read.

2. Advantages and disadvantages of the original and refactored VBA script

*Advantages:*
The structure is clear and easy to understand and read when using.

*Disadvantages:*
Complex to run, slow, not suitable for large amounts of data

