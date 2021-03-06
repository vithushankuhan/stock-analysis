# Green Stocks Analysis

## Overview and Purpose

### Overview
Steve, a recent finance graduate wants to help his parents invest in stocks. More specifically, they want to invest in greeen energy stocks as they believe that as more fossil fuels get used up, there will be an increased demand for alternative energy sources. In this analysis, the performance of a dozen green energy stocks was analyzed. Steve's parents want to invest into DAQO New Energy Corporation. Instead, Steve wants to analyze other green energy stocks in addition to DAQO to get a better picture of the green energy sector, and create a stronger portfolio for his parents. 

### Purpose
In stocks_analysis.xlsm, an analyzis of the dozen stocks was executed using VBA. The Purpose of this challenge is to refactor the original stock_analysis code to make it more efficient for larger datasets. At the end of the analysis, time taken to execute the code was observed. By refactoring the original code, we aim to improve the code's effficiency by taking fewer steps, using less memory, or improving the logic in order to reduce the time taken to execute the code.

## Results

### 2017 vs 2018 Stock Performance
By looking at the images below, it is clear that green stocks performed better in 2017 than in 2018. Looking closely, ENPH and RUN are the only two stocks to consistently be in the green in 2017 and 2018. Every other stock either took a turn and performed worse in 2018 or stayed in the red, like TERP. Steve should look into ENPH and RUN as these stocks could be promising investments. Bot tables can be found in VBA_Challenge.xlsm

<img width="221" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/76541288/110214691-c6686000-7e73-11eb-8b0f-e4c9a219e5c2.png">
<img width="221" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/76541288/110214693-cc5e4100-7e73-11eb-8600-835158c15e09.png">

### 2017 Stock Refactored Code
As seen in the images below, the refactored code significantly improved the efficiency of the analysis on the 2017 performance of green stocks. The refactored code(bottom) was executed 92% quicker than the original code(top). These numbers will vary every time the code is executed, but on average the refactored code is much more efficienct. The refactored message box image can be found in the Resources folder. The orignial code did not contain this variable and as a result had to spend more time looping through

<img width="425" alt="2017 Initial Analysis Time" src="https://user-images.githubusercontent.com/76541288/110214524-fe22d800-7e72-11eb-9de2-3b83d4906495.png">
<img width="423" alt="2017 Refactored Time" src="https://user-images.githubusercontent.com/76541288/110214531-067b1300-7e73-11eb-8d18-c51ad886277e.png">


### 2018 Stock Refactored Code
As seen in the images below, the refactored code significantly improved the efficiency of the analysis on the 2018 performance of green stocks. The refactored code(bottom) was executed 89% quicker than the original code(top).These numbers will vary every time the code is executed, but on average the refactored code is much more efficienct. The refactored message box image can be found in the Resources folder.

<img width="424" alt="2018 Initial Analysis Time" src="https://user-images.githubusercontent.com/76541288/110214555-1dba0080-7e73-11eb-8570-8bd063349542.png">
<img width="424" alt="2018 Refactored Time" src="https://user-images.githubusercontent.com/76541288/110214573-2f030d00-7e73-11eb-8178-397c17de767f.png">

## Summary
In summary, it is clear that refactoring the VBA code can help execute the code in less time, making it more efficient. However, there are other advantages to refactoring code outside of VBA. There are also disadvantages that come with refactoring. 

Advantages
- By Refactoring code, one makes the code easier to read and understand as there might have been a lot of code that may have not been needed. This allows others working on the code be able to understand the thought process behind each line of code. Additionally, refactoring code can help identify bugs. Since the process of refactoring invovles cleaning the code to make it easier to understand, identifying potential bugs in the code becomes much easier. 

Disadvantages
- One of the main disadvantages of refactoring code it the time it takes to refactor code. Since the process of refactoring requires the editor to understand the code thoroughly and identify potential areas of improvement, the whole process of refactoring can take a significant amount of time depending on the length and complexity of the code. To add on to this point, refactoring code for large project might be ineffective. It would be much easier to re-write the code from scratch instead of working out a solution from the initial code. This will also save time which as mentioned earlier can be a major drawback from refactoring code. 

With all that being said, the refactored code for Steve's analysis on green stocks definetly makes it much quicker for Steve to run his analysis. A change made to the refactored code was the reduction of a loop used to access the ticker data. Instead of running two loops to access the tickers, there is now one loop. This allows the data to be accessed much quicker, thus reducing the scripts processing time. In addition to making the code much more efficient, this step also help clean the code to make it easier to understand and read. With fewer steps in the new refactored code, following the structure and thought process of the analysis is much easier for new code readers. A disadvantage however is the approach taken to restructure the code. Instead of re-writing the code from scratch, the original code was used a reference. This could make it more complicated to understand what changes need to be made, and it can also be easy to get lost in all the code. a better approach would be outlining the ways in which the code could be improved and starting from scratch. This would allow for a better understanding of the code and prevent confusion between the original and new code. 


