# **Module 2: VBA of Wall Street**

## **Overview of Project**

### _Purpose_:

The purpose of this project is to create a fast and reliable spreadsheet that Steve can use to conduct research of the stock market.  The original code needed refactoring to handle this new use case and lower the runtime.

### _Results_:

While both codes return results quickly.  The results for the refactored code is consistently faster than the original code.  The original code returned results for [2017](Resources/VBA_Challenge_2017.png) in about 0.559 seconds and completed its analysis of 2018 stocks in about 0.555 seconds [insert picture].  By comparison, the refactored code analyzed 2017 stocks in about 0.121 seconds [insert picture] and returned results for 2018 in about 0.555 seconds [insert picture].

The refactored code runs faster because the inner loop is looping over the array [insert 2 pictures] as opposed to looping through each cell [insert 1 pictures].  Similarly, the output for the “All Stocks” tab is computed using the array [insert 1 photo] as opposed to each cell [insert 1 photo]


## **Summary**

### _Advantages of Code Refactoring_:

Refactored code provides several advantages including a tendency to use less memory and the concise nature of refactored code provides ease of use for others to read and understand the code.  Most importantly refactoring code executes fewer steps, which reduces the probability of human error in writing the code.  These advantages increases the code’s functionality and efficiency

### _Disadvantages of Code Refactoring_:

By way of contrast, refactoring code presents a disadvantage because resources need to be allocated to allow someone re-work code that is complete and functioning.  Therefore, an organizations needs to consider whether the return on investment warrants refactoring code.

### _Pros and Cons of refactoring the original VBA Script_:

Each project is unique.  Therefore, it is necessary to balance the pros and cons before allocating resources to refactoring code.  Refactoring code, on this particular project, makes sense because Steve wants to expand the use of the Excel file beyond its originally intended purpose. The original code was not designed for this new use case so it would have been time consuming to run and there is a possibility that it would fail.  Therefore, refactoring this code is necessary to ensure that the macros work for Steve’s new use case.

