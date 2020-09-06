# Stock analyzing with VBA
Performing analysis on Green_stocks data to uncover which Stock had performed better than others.
## Overview of Project
 Analyze the performance of 12 Green_stocks in 2017 and 2018 based on Total Daily Volume and
 the percentage of yearly return.
 Improve the run_time of code by refactoring it annd getting the same analysis.

## Results

 - Base on the result of the VBA Code in 2017 all of the Green_stocks except TERP 
 had positive yearly return percentage. in the oder words in 2017 many of them 
 performed great.
 - In 2018 only ENOH and RUN stocks had positive return percentage. For all the other
  Green_stocks 2018 was not a good year.
 - In term of executing time, after refactoring the code and use more arrays in compare 
 to the original one, run_time decrease for both years about 0.1 second which is considerably high.
 
 ![VBA Challenge 2017:](Module2/resources/VBA_Challenge_2017.png)
 ![VBA Challenge 2018:](Module2/resources/VBA_Challenge_2018.png)


 ### Summary:
- What are the advantages or disadvantages of refactoring code?
 When refactoring code, you aren’t adding new functionality; you just want 
 to make the code more efficient by taking fewer steps, using less memory, 
 or improving the logic of the code to make it easier for future users to read.

 - How do these pros and cons apply to refactoring the original VBA script?
 For example by using many arrays and use one index for them, we decreased the run-time .

