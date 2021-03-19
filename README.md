# VBA of Wall Street

## Overview of Project
This project is the second weekly challenge of the Data Science Bootcamp. It allows us to put into practice and showcase the skills learned during the second module of the Bootcamp.

### Purpose
The purpose of the project was to provide Steve with a way to quickly generate an analysis of the stock data for a specific year.

By refactoring the macro code generated while working on the module, we are able to generate the output that Steve wanted as well as improve the performance of the code.

## Results
Analysis was carried out in Excel from the source data provided.  Inputs and outputs are located in the Excel file [VBA_Challenge](VBA_Challenge.xlsm).

### Analysis of Stock Performance
As shown in the snapshots below, 2017 was a better year for most stocks in the data, with only one ticker (TERP) showing a negative return.

![All Stocks 2017](/resources/All_Stocks_2017.png)

On 2018, only 2 tickers (ENPH, RUN) had a positive returns, while the rest had negative returns.

![All Stocks 2018](/resources/All_Stocks_2018.png)

Two tickers to look in further detail would be RUN and ENPH.  RUN had a return of +5.5% in 2017 and +84.0% in 2018, making it the only ticker with an increasing return year-on-year.

The ticker ENPH had a return of +129.5% in 2017 and +81.9% in 2018.  Although the return in 2018 is lower than the return in 2017, it total return is still quite significant.

### Analysis of Execution Times
By refactoring and improving on the code created in the module, we were able to reduce the running time to about 1/7th of the original time.

It can easily be imagined that traveling the data set only once (as done in the refactored code) would be an improvement to traveling the entire data set for each of the 12 tickers.

The table below contains the running time from each year for the original code and the refactored code.  Links point to the output running time produced by the code.

| Year | Original Code (s) | Refactored Code (s) | Improvement |
| -----| --------------| --------------- | ------------|
| 2017 | [0.711](resources/Original_code_2017.png) | [0.102](resources/VBA_Challenge_2017.png) | -85.7% |
| 2018 | [0.699](resources/Original_code_2018.png) | [0.106](resources/VBA_Challenge_2017.png) | -84.8% |

Running time was measured from the moment a year was entered until the final formatting of the output table was completed.

### Challenges and Difficulties Encountered
The major challenge encountered was making sure that the original and refactored code were measuring the execution time at the same interval. To do this, I added a call to the formatting subroutine created during the module just before taking the end time value.
```
    Call formatAllStocksAnalysisTable
    endTime = Timer
```

### Future improvements
The code continues to assume that the number and name of the tickers in the data remain constant.  This may not be the case for new years.

The code would benefit from removing the hardwiring the ticker names into the `tickers` array, while maintaining a single pass approach.

## Summary
In this challenge we refactored a VBA macro to improve its performance. 

> Code refactoring is the process of restructuring existing computer code—changing the factoring—without changing its external behavior. [^1]

Among the benefits of code refactoring we find that it:
- improves the design and quality of the code,
- helps finding issues,
- makes the code easier to understand and reuse.

Some of the disadvantages refactoring code are:
- risk of creating bugs when there are no proper test cases,
- risk of creating bugs when the code is not fully understood,
- risk of generating further needs to change in the architecture of the software.

In this challenge code refactoring improved the quality of the code and thus its performance by a factor of 7. It also made the code more legible by the addition of comments.


[^1]: Wikipedia [https://en.wikipedia.org/wiki/Code_refactoring](https://en.wikipedia.org/wiki/Code_refactoring)
