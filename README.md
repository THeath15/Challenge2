# Challenge 2 - VBA Challenge
Refactoring Stock Analysis for Green Stock Energy
## Overview of Project
Steve was satisfied with the workbook prepared for him. To add more to his research for his parents stocks analysis, he wanted to include the entire stock market over the last few years and have a written explanation of the analysis.
### Purpose
 Steve would like  to perform his analysis on larger datasets, and he wants to measure the code performance on how fast the  VBA code will compile the results. To help Steve, we need to add a script that will calculate how long the code takes to execute and output the elapsed time in a message box.

## Analysis and Challenges
In this challenge, we have to edit  or refactor, the Module 2 solution code to loop through all the data in order to collect the same information that we drun through in this module activity.  This will determine if  refactoring  code successfully made the VBA script run faster. Below is the written analysis that explains the findings. 
 
 Using VBA , this project analyzed stock market datasets  for 2 years (2017, 2018 ) The goal is to create a script that will loop through all the stocks for one year for each run and take the following information:

The ticker symbol
Total Daily Volume  and the 
Return or percent change from the starting price  of the given yer to the closig at the end of the year.


## Results

Below are screenshots for the run program results for the refactored VBA Stock Analysis scripts for Steve. 

This screenshot shows the 2017 stock analysis, where ticker for DQ shows highest yield of return at 199.4% and TERP gave a negative return 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/92903447/140622431-33cd940b-d184-4a75-9e6e-7189e587b929.png)

This screenshot shows the 2018 stock analysis, where ticker for DQ shows negative return and ENPH and RUN shows a good percentage of investment 

![VBA_Challenge_2018](https://user-images.githubusercontent.com/92903447/140622434-610f7259-fe03-4cb8-8676-dd2e9d1eb4c6.png)

Run results for 2017:

![Screenshot 2021 VBA Challenge 2017](https://user-images.githubusercontent.com/92903447/140622445-9d178a71-f2b1-4174-937d-a874cc53df28.png)

Run results for 2018 :

![Screenshot 2021-VBA Challenge 2018](https://user-images.githubusercontent.com/92903447/140622450-96a68751-d0a9-4242-b91b-f9368af90abe.png)

The refactored vbscripts shows a faster runtime in providing Steve the information he needs.

![Screenshot 2017 Timer](https://user-images.githubusercontent.com/92903447/140622455-f60c01fc-44cc-4bf1-ae9c-37f2684a8d15.png)

![Screenshot 2018 Timer](https://user-images.githubusercontent.com/92903447/140622459-2dd511d5-5416-4d03-8fa0-5338e12a37a5.png)

- What are the advantages or disadvantages of refactoring code?
Refactoring Advantages:
 - Improves the design  of applications and make software applications more adaptive
 - It helps macros scripts or programming to run faster
 - It helps finding bugs

 Refactor Disadvantages:
 - Time consuming :you might have to retest lots of functionality
 - Chance of mistakes : due to complexity of code , chances of codes to run with more errors

-How do these pros and cons apply to refactoring the original VBA script?
Refactoring makes the code/scripts run more efficiently and improves the logic of the original code but could be risky and can take extra time and budget if the developer do not know the original code and could introduced more bugs.
