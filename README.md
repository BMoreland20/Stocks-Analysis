# Stocks-Analysis
Overview of Project:
	In this analysis we are observing and visualizing the growth and decay of 12 stocks over a year long period for 2017 and 2018.  In this we show the total daily volume of each stock and the return whether it be positive or negative in value.


Results:
	In 2017 all but one stock came back with positive returns, the stock that failed to yield was TERP. ![This is an image](https://github.com/BMoreland20/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2017_Results.png)  Where as in 2018 all but two failed to make a positive return those being ENPH and RUN. ![This is an image](https://github.com/BMoreland20/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2018_Results.png)  To perform this analysis we first have to create code to automatically perform the required analysis quicker than we could manually.  After setting up the array and listing all of the variable we need to build the 'For' loop and multiple 'If' 'Then' statements for pull from the total volume and look at the starting and ending prices to determine our return.  Here is some examples of the code needed to run these loops:
  
`For j = 2 To RowCount`

`If Cells(j, 1).Value = ticker Then  
totalVolume = totalVolume + Cells(j, 8).Value 
End If`

`If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then 
startingPrice = Cells(j, 6).Value
End If`

`If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
 	endingPrice = Cells(j, 6).Value
End If
Next j`  

These lines of code allow us to automate the process of analyzing the returns of the stocks.  With this addition of formatting code we can quickly format the results to show which had a positive yield vs a negative yield.  However, this code in total ran slowly at about .6 seconds for both years (see images). ![This is an image](https://github.com/BMoreland20/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2017_Original.png) 
![This is an image](https://github.com/BMoreland20/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2018_Original.png)

Now if we refactor the middle section of the code we can speed up the time it takes to perform this analysis (see images). ![This is an image](https://github.com/BMoreland20/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2017.png) 
![This is an image](https://github.com/BMoreland20/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2018.png)
This refactoring saved some time but not much by bringing the time down to roughly .55 seconds for each respectively.

To make this change here is some of the code that was changed to speed the code up.

`tickerIndex = 0`

`Dim tickerVolumes(12) As Long`
`Dim tickerStartingPrices(12) As Single`
`Dim tickerEndingPrices(12) As Single`
    
    `For i = 0 To 11
        tickerVolumes(i) = 0
    Next I`


Summary:
	
Question 1): One of the main advantages of refactoring code is that it can speed up the time it takes to run said code.  If you have a large data set these small changes could net seconds if not minutes off of the time it takes to compile the results.  This in mind a disadvantage is that you could spend hours to have a small return in time it takes to compile the results.  The other disadvantage I can think of is that you could end up “breaking” your code and have to troubleshoot and debug these new lines of code.
	
Question 2): As mentioned above the time it could take to refactor the code could take more time than it would to leave the code as is.  As we saw with this example both codes were originally running at about .6 seconds but after refactoring the code took about .55 seconds to complete.  To have .05 hundredths of a second saved is not detectable by a human unless you have a timer running as we had.
But if you are working with a very large data set you could potentially save minutes if not hours of compiling time when you refactor the code.
Balancing the pros and cons that I listed I would say that it would be on a conditional basis if it is worth refactoring the code.  If you’re looking at milliseconds in the thousandths place it may not be worth refactoring the code.  However, if you’re looking at code that takes about for example 30 minutes to compile it would be worth taking a look at your code to refactor it.
While confirming how to best refactor my code I found a YouTube video that showed breaking down your code into parts to best determine where you would be best spending your time looking to improve your code.   [YouTube](https://www.youtube.com/watch?v=RNqd89K_bbU&ab_channel=ExcelMacroMastery)
