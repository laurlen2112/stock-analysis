# **VBA of Wall Street**

## **Overview of Project**

### _Purpose_:

The purpose of this project is to create a fast and reliable spreadsheet that enables Steve to conduct research of the entire stock market.  The original code needed refactoring to handle this new use case and decrease the computing runtime.

### _Results_:

While both codes return results quickly, the result for the refactored code is consistently faster than the original code.  The original code returned results for [2017](Resources/Original_Code_Runtime_2017.png) in about 0.559 seconds and completed its analysis of [2018](https://github.com/laurlen2112/stock-analysis/blob/main/Resources/Original_Code_Runtime%202018.png)  stocks in about 0.555 seconds.  By comparison, the refactored code analyzed [2017](Resources/VBA_Challenge_2017.png) stocks in about 0.121 seconds and returned results for [2018](Resources/VBA_Challenge_2018.png) in about 0.125 seconds.

The refactored code runs faster because the inner loop is looping over the arrays (pictured [here](Resources/Refactored_Code-2B-3b.png) and [here](Resources/Refactored_3b_to_3d.png)).  Similarly, the [output](Resources/Refactored_Code_4.png) of the refactored code is computed using the arrays.  By contrast, the [original code](https://github.com/laurlen2112/stock-analysis/blob/main/Resources/Original_Code%204%20to%205c.png) loops through each cell and its [output](Resources/Origina_Code_5d_to_7.png) is computed by cell.

## **Summary**

### _Advantages of Code Refactoring_:

Refactored code provides several advantages such as a tendency to use less memory.  Additionally, the concise nature of refactored code provides ease of use for others in reading and understanding it.  Most importantly, refactoring code executes in fewer steps, which reduces the probability of human error in writing the code.  These advantages increase the code’s functionality and efficiency

### _Disadvantages of Code Refactoring_:

By way of contrast, refactoring code presents a disadvantage because resources must be allocated to allow someone re-work code that is complete and functioning.  Therefore, organizations need to consider whether the return on investment warrants refactoring code.

### _Pros and Cons of refactoring the original VBA Script_:

Each project is unique.  Thus, it is necessary to balance the pros and cons before allocating resources to refactoring code.  Refactoring code, on this particular project, makes sense because Steve wants to expand the use of the Excel file beyond its original purpose. The original code was not designed for this new use case so it would have been time consuming to run and there is a possibility that it would fail.  Therefore, refactoring this code is necessary to ensure that the macros work for Steve’s new use case.

