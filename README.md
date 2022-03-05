# Stock-Analysis Using VBA
VBA of Wall Street

# Overview
Assist Steve to help his parents analyze the performance of green every stocks in the years 2017 and 2018. 

## Purpose
Conduct an initial stock analysis for Steve and refractor the original VBA code to observe if the code will run more efficiently by measuring the time it takes to complete the code (original VBA script vs refractory script).

The VBA code uses a timer, arrays, if-then conditional statements, assigns long/string data types, and adds static/conditional formatting.

# Results
In order to make the code run more efficiently, I needed to switch the nesting order of the `for` loops. To achieve this, I first created a variable called the `tickerIndex` that accessed the correct index across the `tickers` array as well as the 3 new output arrays I created next: `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`. Setting up the `tickerIndex` variable meant the code was able to assign the `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` to each ticker symbol before looping through the entire dataset. Refactoring the code in this manner allowed it to run much faster than using the original script.
 
## Comparing the Original Run Times to the Refactored Run Times

Run times for the original code took around .71 seconds for 2017 and .64 seconds for 2018.

![Resources/Code Performance_Module(2017)](Resources/Code Performance_Module(2017).png) 

![Resources/Code Performance_Module(2018)] (/Resources/Code Performance_Module(2018).png)

Run times for the refactored code took around .10 seconds for 2017 and 2018

![Resources/VBA_Challenge_2017] (/Resources/VBA_Challenge_2017.png) 

![Resources/VBA_Challenge_2018] (/Resources/VBA_Challenge_2018.png) 

Refactoring the code did make the run times decrease, which optimizes the code. 



# Summary
## 1. What are the advantages or disadvantages of refactoring code?
### Advantages
Finding the root cause of a potential bug can be done with refactoring.  A programmer can catch duplicated subroutines, unnecessary loops, redundant statements, or code that was used to run down an error but was accidently left in the script.  Another advantage is when a peer reviews another programmer’s work, they will come with a fresh perspective and can tidy up the code to make it run leaner. 

### Disadvantages

Programming can have multiple approaches to a solve a problem.  In those differences, programmers may have alternative logic steps, which will require testing to see how those differences play out in the script.  Refracting a stable code to apply a different set of logic could be costly or introduce new bugs into the system.  When juggling tight deadlines, programmers may also have to choose between refactoring or developing new code.  
	
## 2. How do these pros and cons apply to refactoring the original VBA script?
Reducing the number of loops decreases the memory needed for processing the data, which reduces the run time and optimizes the performance of the script. To refactor the code, testing has to be done with each new addition to check for the efficiency of the new code.  
