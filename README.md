# ahoynacke-CWRUBootcamp_M2_081622_StocksAnalysis_Hoyncke
Stock analysis using VBA 


# VBA Challenge 

## Overview of Project

### Purpose 
    The purpose of week 2 VBA Challenge is to refactor our original code created to collect stock information relating to certain green/eco-friendly companies during 2017 and 2018. 

### Background
    After analyzing green oriented stocks for Steve, I now want to expand on the dataset to the entire stock market over the years 2017 and 2018. Since this will return a large amount of data, we need to create efficient and concise code by refactoring our original code.
    
## Results: 

### Compare performance of 2017 & 2018. 
    The data returned by the code includes a table on 12 different stocks. The code prompts the user to identity which year they would like they would like the data ran for. The information returned in the table includes the ticker Value or stock name, total daily volume and the return percent. I started the process of improving the codes performance and execution time by copying the code provided. I then edited and refactored the code to loop through all the data one time to collect the requested information. 

      ''2a) Create a for loop to initialize the ticker Volumes to zero.
'    Did that by adding () above
'    Other option:
       For I = 0 To 11
       ticker Volumes(I) = 0
'       tickerStartingPrices (I)
'       tickerEndingPrices (I)
           Next I

    ''2b) Loop over all the rows in the spreadsheet.
    For I = 2 To RowCount

        '3a) Increase volume for current ticker
        ticker Volumes(tickerIndex) = ticker Volumes(tickerIndex) + Cells(I, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(I - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(I, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(I, 6).Value
                tickerIndex = tickerIndex + 1
            
            End If
    

### Compare Execution times of the original script and the refactored script.
    Below are the execution times from the original script (green stocks) and the refactored script (VBA challenge)
    green stocks 2017 .304
    green stocks 2018 .296
    VBA Challenge 2017 .089
    VBA Challenge 2018 .0625
    Significant decrease in run time between the original and refactored code. 
## Summary 

### What are the advantages or disadvantages of refactoring code?
    The advantaged of refactoring code include cleaner, easily digestible, and organized code. In addition, the code typically runs faster because by refactoring we are making the code more concise and efficient. 
    Disadvantages include the time and resources spent cleaning up code and it is a risky activity if the code is large and very intricate. 

### How do these pros and cons apply to refactoring the original VBA script?
    In this case the code was fairly simple to begin with so there was little risk. We were able to refactor the code without breaking the entire script and get a faster return rate for the data 
