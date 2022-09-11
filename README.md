# stocks-analysis
OVERVIEW: VBA Stock Analysis Project

Purpose
The project looks at editing, or refactoring, the Stock Market Dataset with VBA solution code so that it loops through the selected  data once in order to collect information from an entire dataset. Furthermore, we decide if refactoring the code made the VBA script run faster. Then, we decide how to make the code more efficient through a condensed model which uses less memory and makes it easier for others to read.

Analysis and Challenges
A quick look at the Kickstarting Analysis and Challenges of this Project will show the following tasks:

Prepare our dataset VBA_Challenge.vbs file for the project.
Create our resources folder in GitHub to hold the run-time pop-up messages that we’ll screenshot after running refactored analyses for 2017 and 2018.
Create and convert our XLSM file from *.vbs dataset that you used in this module as VBA_Challenge.xlsm.
Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
Use the steps Refactor VBA code and measure performance to add code where indicated by the numbered comments in the starter code file.
Use your knowledge of VBA and the starter code provided in this Project to refactor the VBA Script dataset so we loop through the data one time and collect all of the information.

Our Challenge Data Background
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.
This a file covering the basics of how to perform stock analysis of a data sheet from 2017-2018

RESULTS: Refactor VBA Code and Measure Performance
Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:
1. The tickerIndex is set equal to zero before looping over the rows.

Created a tickerIndex variable and set it equal to zero before iterating over all the rows. Will use this tickerIndex to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requierement.

![arrays](https://user-images.githubusercontent.com/111712209/189547263-b95842b1-226d-4bab-97c0-6a2d6b1a987d.png)


2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. In our VBA code, the tickerVolumes array should be a Long data type. But in our VBA code the tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

![tickerVolumeStartingPriceEndingPrice](https://user-images.githubusercontent.com/111712209/189547425-6984c055-a58e-4d03-96ce-70fb2948262c.png)


3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

Created a for loop to initialize the tickerVolumes to zero. And if the next row’s ticker doesn’t match, increase the tickerIndex.

![Samplecode](https://user-images.githubusercontent.com/111712209/189547369-47804552-c810-4766-aa1d-455eec30e68e.png)


4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

![ForLoop](https://user-images.githubusercontent.com/111712209/189547457-c7f3a9fb-10e6-44b9-8f6e-3f05a4c1dad6.png)


Stored values from tickerStartingPrices and tickerEndingPrices

Created an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current closing price to the tickerStartingPrices and tickerEndingPrices variable.

![If-thenStatement](https://user-images.githubusercontent.com/111712209/189547605-ea7226d3-8ac1-4efb-b684-d2df4d982da1.png)


5. Code for formatting the cells in the spreadsheet is working.

Positive returns are displayed in green while negative returns are shown in red. This makes it easier to determine which stocks did performed well based on formatting done on the values of the return.

6. There are comments to explain the purpose of the code.

Adding Comments is required, as a Best Practices for Writing Super Readable Code such as for,

Commenting & Documentation,
Consistent Indentation,
Avoiding Obvious Comments.
Code Grouping,
Consistent Naming Scheme,
DRY (Don't Repeat Yourself) Principle,

7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module

Finally, we run the stock analysis, to confirm that our stock analysis outputs for 2017 and 2018 are the same as dataset example provided (as shown in the images below, named Dataset Examples Provided). In adition, in our resources folder and below you can see the final Stock Analysis Results named, Final VBA Analysis 2017 and 2018 save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. Then, save the changes to your workbook..

Dataset Examples Provided


![2017-DataSet](https://user-images.githubusercontent.com/111712209/189547670-75d0a9b2-92f9-44fc-bc83-d243a5d000f1.png)
![2018-DataSet](https://user-images.githubusercontent.com/111712209/189547675-20c2dd17-0c7d-47c8-993f-1c7b215342dc.png)

Below our Final VBA Analysis PNGs,

Final VBA Analysis 2017

![VBA_Challenge_2017](https://user-images.githubusercontent.com/111712209/189547709-f2366287-4a62-472e-9bac-de8aaf55751b.png)


Final VBA Analysis 2018

![VBA_Challenge_2018](https://user-images.githubusercontent.com/111712209/189547688-dd3bd2a4-5021-42e4-acd4-95ee5ace6000.png)


8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png

Running our fully 2017 and 2018 data stock analysis gave us an elapsed run time for each year, below our results.

Time on VBA_Challenge_2017.PNG

![VBA_Challenge_2017](https://user-images.githubusercontent.com/111712209/189547764-6b580306-f5d1-4390-a474-b39399c71db0.jpg)


Time on VBA_Challenge_2018.PNG

![VBA_Challenge_2018_](https://user-images.githubusercontent.com/111712209/189547770-b4a041db-acbb-4835-bcb9-447bbaa51ffd.png)


SUMMARY: Our Statement:
Deliverable with detail analysis:
1. What are the advantages or disadvantages of refactoring code?

You need to perform code refactoring in small steps. Make tiny changes in your program as you work through it, each of the small changes improves the efficiency of your code, leaving the application updated.

Disadvantages:

A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
A complex unstructured code is usually best to split in several functions.

Refactoring process can affect the testing outcomes.

Advantages:

Logical errors can appear in a well structured code document that contains nested, conditionals and loops.
In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.

2. How do these pros and cons apply to refactoring the original VBA script?

Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring. In the instance that one needs to troubleshoot there code, they will have a better understanding since they restructured or improved their code. 

By writing clean code in the refactoring process, we are reducing clutter which makes adding changes alot easier. This allows you to maintain your code in an organized manner so that its easy to understand and avoid issues on later
