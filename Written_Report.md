# **Module 2 Challenge**

## **Overview of Project**
   The main overview of the project is to help Steve to expand his dataset to include the entire stock market over the last few years.
      

### **Purpose**

The main pupose of the project is to refactor the Module 2 solution code to loop through all the data in once and collect the same information as we did in Module2 
and verify whether the refactored code made the VBA script run faster.



## **Results**
Completed all the requirements as per Deliverable1 
   - Created the tickerIndex and intialized to 0.
   - Created all the respective arrays for tickers,tickerStartingPrices,tickerEndingPrices.
   - The script looped through all of them respectively.
   - code for formatting the cells in spreadsheet is working as expected.
   - All the comments are embedded within the modules in respective places.
   - The outputs of the non refactored code and the refactored code exactly.
   - The timer time based of start and ending time has been captured and uploaded
   


#### **Stock Analysis of year 2018 without refactoring**
![img](https://github.com/hsurisetti/StockAnalysis_Challenge/blob/main/resources/VBA_Challenge_2018.png)
#### **Stock Analysis of year 2018 with refactoring**
![img](https://github.com/hsurisetti/StockAnalysis_Challenge/blob/main/resources/VBA_Challenge_2018_refactored.png)



## **Summary**  
  In Summary , it is a good coding practice to always keep refactoring the code, to ensure a good software quality.

- **What are the advantages and disadvantages of refactoring code?**

   After refactoring, the code looks fresh and easier to read and understand.
   It is less complex and easier to maintain thereby improving the quality of the software.

   Some of the disadvantages of refactoring involve, consuming too much time in refactoring as many things might need to sorted out first, even for refactoring a small piece of code. We also need to make sure that the refactored code doesn't change any of the exisitng functionalites which otherwise would lead to regression.

- **How do these pros and cons apply to refactoring the original VBA script ?**

    The pros of refactoring this VBA script invlove , cleaner looking code with better readability
    The non refactored code had to activate the speadsheets multiple times in the loop to output the data. With the refactored code, we are iterating over the spreadhseet and collecting all the outputs like volumes,starting and ending prices in their arrays and 
    accessing them in the end to dispaly the output in "All Stock Analysis" sheet.
    The time taken to run the code using the refactored code is far less than the non refactored code thereby improving the performance of the software tremendously, especially when a large data set is invloved.

    Some of the cons of the refactored code, although not much in this case is , it takes some extra time to construct new arrays and use them efficiently without going over multiple loops. 

