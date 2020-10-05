# stock-analysis
An analysis of several Green Energy stocks with respect to their yearly volumes and overall return. 

## Overview of Project and Purpose
The purpose of this project is to provide a client with an analysis of Green Energy stock data over the course of 2017 and 2018. The initial data set is lengthy and contains a ticker for every stock and its daily records including opening cost, highs, lows, closing cost, and daily volume for each year. Since there are 3000+ rows of data and 8 columns, it is helpful to use a tool like Excel’s VBA to write macros to analyze the data all at once. Writing a code through VBA also allows the data to be updated or have additional years added for further analysis rather than the sometimes time consuming process of manually analyzing. The main objective of the challenge was to refactor the code in VBA that was already written to loop through all the data just one time in order to output the ticker symbol, total daily volume, return, and apply formatting for each stock within the year. 

## Results
After running the stock analysis VBA script, 2 tables were generated for 2017 and 2018 containing the ticker data in relation to their total daily volumes and yearly return. In 2017, one stock with the ticker *TERP* had a negative return while all other stocks yielded a positive return. In 2018, only 2 stocks had a positive return for the year: *ENPH* and *RUN*. It would not be a good investment to put all your funds into *DQ* like the original analysis proposed, because although they had a positive return in 2017, the total daily volume traded was significantly lower than the following year that yielded a -62.6% return. On the other hand, *ENPH* had a positive return for both 2017 and 2018. It also significantly increased its total daily volumes from 221,772,100 to 607,473,500 which would make it a good stock to consider investing in.  See attached tables below for 2017 and 2018 outputs. 

![alt text]( https://github.com/coconnell022/stock-analysis/blob/main/All%20Stocks%202017.png?raw=true)
![alt text]( https://github.com/coconnell022/stock-analysis/blob/main/All%20Stocks%202018.png?raw=true)

The original script contained an execution time of around 0.63 seconds for 2017 and 2018 while the refactored code contained an execution time of around 0.12 for both years. It can be clearly seen that refactoring the code allows the analysis to be completed quicker. See attached images below. 

  - Original script run times:

    ![alt text]( https://github.com/coconnell022/stock-analysis/blob/main/Original%20Script%202017.png?raw=true)
    ![alt text]( https://github.com/coconnell022/stock-analysis/blob/main/Original%20Script%202018.png?raw=true)

  - Refactored script run times:

    ![alt text]( https://github.com/coconnell022/stock-analysis/blob/4dd3d7f7d0382061d801ebdc17d92c2e94badf42/VBA_Challenge_2017.png?raw=true)
    ![alt text]( https://github.com/coconnell022/stock-analysis/blob/4dd3d7f7d0382061d801ebdc17d92c2e94badf42/VBA_Challenge_2018.png?raw=true)

Refactoring the code took away the need for nested loops to be included within the macro which sped up the execution time and cleaned up the lines of code. See below for a section of the refactored code that contains 3 seperate "for" loops. 

```
 '1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
            
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
        For i = 0 To 11
    
            tickerVolumes(i) = 0
            
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
             '3a) Increase volume for current ticker
             
                     tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                 
             '3b) Check if the current row is the first row with the selected tickerIndex.
            
                 If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                     tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                 End If
                 
             '3c) check if the current row is the last row with the selected ticker
            
                 If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                     tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                     
             '3d) Increase the tickerIndex.
             
                     tickerIndex = tickerIndex + 1
                         
                 End If
                                   
       Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
            Worksheets("All Stocks Analysis").Activate
            
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
                
        Next i
```

## Summary
1.	What are the advantages or disadvantages of refactoring code?

- Refactoring VBA code allows the analysis to operate more efficiently by simplifying it. This refactoring process makes the code more understandable for new users. The new optimization of the code does not change the output behavior and improves the overall quality of the code by restructuring potentially confusing codes. By introducing more arrays, refactoring also allows for the code to be run faster so if there is a specific time frame in which multiple data sets need to be analyzed this could prove to be highly beneficial. However, the refactoring process can be time consuming and often cause errors that need to be debugged. It can also be risky to perform refactoring on an already working or stable code, even if that code is lengthy. Overall, if time permits, it would be favorable to refactor most code so it can be used for future reference with less difficulties. 

2.	How do these pros and cons apply to refactoring the original VBA script?

- In this specific stock analysis, refactoring did not provide much clarity in the overall code structure. Although the amount of time it takes to run the code was slightly reduced from 0.63 seconds to 0.12, I do not think that refactoring benefited this analysis. Unfortunately, after refactoring, there are still the same number of lines within the code and it is not necessarily easier for users to understand. In the future if more years of data were to be added or stocks outside of Green Energy were to be analyzed, then refactoring would be worth it. 

