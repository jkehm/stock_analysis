# Stock Analysis Module 2 Challenge

## Overview

### The Data
The data that we will be working with was provided. It consists of two worksheets titled "2017" and "2018". These sheets contain the stock information with a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. Our goal is to utilize macros in order to output a table with the 12 different ticker values, as well as the Total Daily Volume, and Return on the stock for that year. This will allow Steve's parents to make a more informed decision when looking at these stock options.  

### The Purpose
We have already helped our friend Steve create an Excel workbook with Macros utilizing VBA (Visual Basic Analysis) programming. The Macro in class works pretty well for a few stocks, however it would struggle with a larger dataset. In this assignment we will refactor the code we had already written. Ideally, the macro will run much more efficently and take significantly less time to run.

## Results

![VBA 2017 Screenshot](https://github.com/jkehm/stock_analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_2018_Screenshot](https://github.com/jkehm/stock_analysis/blob/main/Resources/VBA_Challenge_2018.png)



## Summary
###   1) What are the advantages or disadvantages of refactoring code?
#### Once a script or macro is written and works succesfully, some people may think their work is done and it is time to move on to the next project. However, there are some reasons to go back and refactor your code to work better. Of course, there are some reasons and situations where it may **not** be worth taking the extra time to do this. 
#### The most obvious advantage to going back to refactor code could be to decrease the run time of the script. This may not seem like a big deal on a fairly small data set like we used here. However, as datasets get larger you could be saving several minutes in run time if the script is written more efficently. Another benefit is to go back and add comments, or add more detail to existing comments. This way, if the script needs to be changed you still understand what everything is doing, even if it is years later. Or if someone else were to pick up the project they can make sense of the code.
#### The biggest disadvantage is that the process of refactoring can be very time consuming. And in a company setting, it may not be worth it to invest more time, energy, and money into a project if it is already satisfactory. Theoretically refactoring a code that already works may lead to bugs in the code as well, if precautions are not taken. 

###   2) How do these pros and cons apply to refactoring the original VBA Script?
#### In this example the main benefit that was clearly noticed was the efficency of the new script. Where the original Script took right about 1 second to run, the refactored version takes about 0.1 seconds to run. The comments on this script were also more step-by-step and easier to follow for someone else to understand what each piece of code is doing.
#### The first disadvantage I mentioned above would not apply to this situation at all. Since we are still learning, refactoring this code was a very useful exercise and we are not costing a company anything by continuing to work on this script. Causing bugs or having run issues is a valid disadvantage. But again, we are learning here, so the process of debugging is extremely important to get thorough and learn from. 
