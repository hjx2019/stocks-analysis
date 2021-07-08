# stocks-analysis
## Project Overview
> Purpose of this analysis

In this project, the stock data is analyzed. Using the VBA Loop and Condition code, Total Volume and Return Rates can be calculated. Using conditional formatting, return rates are shown in colors(Green/Red). Using Index in the Loop, monitored running time is shortened.

## Project Results
1. **Input Box**

Use input box to get the year of data to be analyzed.

![Input box](/Resources/Inputbox.png)

> yearValue = InputBox("What year would you like to run the analysis on?")

2. **Result of 2017/2018 before the refactoring**

On my PC, the run time is around **0.8s before** the refactoring.

![Result 2017 before](/Resources/Run-time2017.png)

![Result 2018 before](/Resources/Run-time2018.png)



3. **Refactoring**

> In the old version, for each ticker in the outer loop, 'IF' checked through all rows of the data.

![Code1](/Resources/refactoringLoop.png)

> The refactoring idea is to mark the ticker during the Loop. When the order of the Tickers Array is the same as the data, only one loop is needed to get the result.

![Code2](/Resources/NewLoop.png)

4. **Check the new run-time**

> New run-time is about **0.14s**, which is significantly faster than **0.8s**. Refactoring succeed!

![Result 2017 after](/Resources/Run-Time-Refactored2017.png)

![Result 2018 after](/Resources/Run-Time-Refactored2018.png)


## Summary

### What are the advantages or disadvantages of refactoring code?

**Advantages**
Formatting is familiar to users, no need to re-design.
Just implement the required functions and don't need to touch other parts of the code.

**Disadvantages**
Structures are established, so it is complicated if some logical relationship needs to be changed. Accordingly, the structures can not be wrong.


### How do these pros and cons apply to refactoring the original VBA script?
**pros**
Just modified the looping part, then the job is basically done.
The new method reduced the loading of looping and the accuracy of data is kept.

**cons**
The logic of this coding is good, so there are no big cons in this work.

**At the end**
Hard-code of assigning the Tickers Array value also needs to be optimized.
Data must be clean (well organized): grouped by tickers and ordered by date. That's a limitation. Also needs to be automized.
