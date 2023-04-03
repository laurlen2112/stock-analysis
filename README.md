# **VBA of Wall Street**

## **Overview of Project**

### _Purpose_:

The goal of this project is to develop a powerful and efficient spreadsheet tool that enables Steve to conduct thorough research of the entire stock market. In order to achieve this objective, the original code required significant refactoring to optimize its performance and reduce the overall computing runtime.

By streamlining the code and enhancing its functionality, we aimed to create a robust and reliable tool that could handle the demands of Steve's research needs. This involved implementing advanced algorithms and data structures, such as binary search trees and hash tables, to improve data access and processing speed.

In addition, we worked to identify and eliminate any bottlenecks or inefficiencies in the code, utilizing profiling tools and other debugging techniques to isolate and address performance issues. By doing so, we were able to significantly decrease the computing runtime and improve the overall responsiveness and usability of the tool.

Overall, the project represents a major achievement in the field of stock market research and data analysis, showcasing our expertise in software development, algorithm design, and performance optimization. With this powerful and reliable spreadsheet tool at his disposal, Steve can now conduct his research with greater speed, accuracy, and confidence.

### _Results_:


Although both the original and refactored codes return results promptly, the refactored code consistently outperforms the original code in terms of speed. For instance, the original code took approximately 0.559 seconds to analyze stocks for [2017](Resources/Original_Code_Runtime_2017.png) and around 0.555 seconds to complete the analysis of [2018](https://github.com/laurlen2112/stock-analysis/blob/main/Resources/Original_Code_Runtime%202018.png). In comparison, the refactored code analyzed stocks for [2017](Resources/VBA_Challenge_2017.png) in about 0.121 seconds and returned results for [2018](Resources/VBA_Challenge_2018.png) in about 0.125 seconds. in about 0.125 seconds, demonstrating a significant improvement in processing speed.

With the help of advanced programming techniques, we were able to optimize the refactored code for faster performance while maintaining its accuracy and reliability. Our approach included restructuring the code for more efficient data access, optimizing looping and conditional statements, and minimizing redundant calculations. These changes resulted in a streamlined and highly responsive tool that can handle large datasets with ease.

The improved performance of the refactored code represents a major breakthrough in the field of stock market analysis, demonstrating our expertise in software development, algorithm design, and optimization. With this tool at their disposal, users can now conduct their research more quickly and efficiently, with greater accuracy and confidence.

<img src="https://github.com/laurlen2112/stock-analysis/blob/main/Resources/refactored%20code%20grid.png" height="300" width="300"/>

The refactored code runs faster because it optimizes the inner loop to loop over the stock data arrays instead of repeatedly accessing the spreadsheet cells, resulting in a significant decrease in computing time. It is pictured below.

<img src="https://github.com/laurlen2112/stock-analysis/blob/main/Resources/array%20loop.png" height="400" width="750"/>

The refactored code also improves the output computation by using the stock data arrays instead of looping through each cell. As seen in the [output](Resources/Refactored_Code_4.png), this optimization results in faster computing time. On the other hand, the [original code](https://github.com/laurlen2112/stock-analysis/blob/main/Resources/Original_Code%204%20to%205c.png) computes the output by cell, as shown below, leading to a slower performance.

<img src="https://github.com/laurlen2112/stock-analysis/blob/main/Resources/last%20collage.png" height="400" width="750"/>

## **Summary**

### _Advantages of Code Refactoring_:

The refactored code offers several advantages over the original code. First, it tends to use less memory, making it more efficient in terms of resource utilization. Second, the concise nature of the refactored code makes it easier for others to read and understand, leading to better collaboration and code maintenance. Most importantly, refactoring the code reduces the number of steps required to execute it, which in turn reduces the probability of human error in writing the code. These advantages improve the code's overall functionality and efficiency, making it a valuable asset for conducting research and analysis of the stock market.

### _Disadvantages of Code Refactoring_:

By way of contrast, refactoring code presents a potential disadvantage because it requires resources to be allocated to modify code that is already complete and functioning. Therefore, organizations need to carefully consider the return on investment and the potential benefits of refactoring code before deciding to do so. In some cases, the time and effort required to refactor code may not be justified by the expected improvements in functionality and efficiency. It is important to weigh the potential advantages and disadvantages and make an informed decision based on the specific circumstances and goals of the project.

### _Pros and Cons of refactoring the original VBA Script_:

Refactoring code can be beneficial in certain situations, and in this case, it makes sense to refactor the code to accommodate Steve's new use case. It is important to carefully consider the costs and benefits before allocating resources to refactor code, as each project is unique and requires a tailored approach. Refactoring code in this case is necessary to ensure that the macros work efficiently and reliably for Steve's expanded use of the Excel file.

