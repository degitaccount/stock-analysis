# **Overview**
The purpose of this analysis is to gather total volume and percentage return for twelve individual stocks over a one-year period, at one time.  For this analysis I refactored existing VBA code to loop through the data and produce the output.  After successfully refactoring the code, I tested the speed and made further edits to make it more efficient.

### **Results**
*VBA Code*

1a) Create a ticker Index variable and set it equal to zero before iterating all the rows

    for i = 0 To 11
    tickerIndex = tickers(i)
        
1b) Create three output arrays: tickerVolumes, tickerStartingPrices and tickerVolumes

    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
    
2a) Create a for loop to initialize the tickerVolumes to zero

    Worksheets(yearValue).Activate
    totalVolumes = 0
                
2b) Loop over all the rows in the spreadsheet
    
    Worksheets(yearValue).Activate
    For k = 2 To RowCount
        
3a) Increase volume for current ticker 

    If Cells(k, 1).Value = tickerIndex Then
    totalVolumes = totalVolumes + Cells(k, 8).Value
    End If
        
3b) Check if the current row is the first row with the selected tickerIndex
        	
    If Cells(k, 1).Value = tickerIndex And Cells(k - 1, 1).Value <> tickerIndex Then
    tickerStartingPrices = Cells(k, 6).Value
    End If

3c) Check if the current row is the last row with the selected ticker, If the next row’s ticker doesn’t match, increase the tickerIndex
            
    If Cells(k, 1).Value = tickerIndex And Cells(k + 1, 1).Value <> tickerIndex Then
    tickerEndingPrices = Cells(k, 6).Value
    End If

3d) Increase the tickerIndex and Loop through the arrays to output the Ticker, Total Daily Volume, and Return
      
Next k
     
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickerIndex
    Cells(4 + i, 2).Value = totalVolumes
    Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1

Next i

*Output and Code Performance*

The analysis shows that the overall performance of the stock porfolio was far superior in 2017, in comparison to 2018.  Only one stock, TERP, had a negative return in 2018 while one-third (4 out of 12) delivered returns of 100% or more during the year.  In 2018, 83% of the stocks had negative returns while ENPH and RUN continued to outperform with returns of 81.92% and 83.95%, respectively.  While TERP delivered negative performance in both years the decline was lower in 2018 incomparison to 2017.

![VBA_Challenge_AllStocks2017](https://github.com/degitaccount/stock-analysis/blob/main/Resources/VBA_Challenge_AllStocks2017.png)    ![VBA_Challenge_AllStocks2018](https://github.com/degitaccount/stock-analysis/blob/main/Resources/VBA_Challenge_AllStocks2018.png)

The final refactored code was far more effecient as it cut the processing time roughly in half.

![VBA_Challenge_2017](https://github.com/degitaccount/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)   ![VBA_Challenge_2018](https://github.com/degitaccount/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

### **Summary**

Refactoring code has several *advantages* but also some *disadvantages*:

| Advantages                                            | Disadvantages                                                                                         | 
| :---------------------------------------------------- | :-----------------------------------------------------------------------------------------------------| 
| Can take less time than writing the code from scratch | Poorly written code without comments may be difficult to follow                                       | 
| Code will be better organized                         | When using someone else’s code you may unknowingly download something malicious                       |
| Code could run more efficiently                       | It will not change the output code that is already operable and therefore may not always be necessary | 

Refactoring the code for this project did not change the output, however it did improve the speed of the analysis.
