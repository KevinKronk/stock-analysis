# Stock Analysis with VBA

## Overview of the Project

### Purpose

## Results

### Stock Performance: 2017 vs 2018

![Stock Returns 2017](Resources/Stock_Returns_2017.png)
![Stock Returns 2018](Resources/Stock_Returns_2018.png)

### Execution Times: Original vs Refactored Script

Original Macro Execution Times

![Original Macro 2017](Resources/Original_Stock_Macros_2017.png)
![Original Macro 2018](Resources/Original_Stock_Macro_2018.png)

Refactored Macro Execution Times

![Refactored Macro 2017](Resources/VBA_Challenge_2017.png)
![Refactored Macro 2018](Resources/VBA_Challenge_2018.png)

## Summary
- Advantages or disadvantages of refactoring code

Refactored code can be more efficient, which can save a lot of time and money in a business setting, as well as improve an existing product. Additionally, refactored code can be more readable, making it easier to understand what's being done. The downsides are that it can take time to refactor the code, and it is essential that the code is tested and works before replacing the old code. One final disadvantage could be that people who were used to the old code will need to reaquaint themselves with the new code.

- How do these pros and cons apply to refactoring the original VBA script? 

The refactored macro ran more than 6x faster than the original macro. In my opinion the code is also easier to understand as you only go through the data once adding values to arrays, and then you print those values with formatting afterwards. There were some initial issues getting the return values to be correct (needed to use closing price on both the Starting and Ending values), once fixed, the output was correct. While it did take time to refactor the original VBA script, if this was a process that was run every day on a much larger file, then the increase in speed would be worth it.
