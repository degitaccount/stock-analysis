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
    
2a) Create a for loop to initialize the tickerVolumes to zero.

    Worksheets(yearValue).Activate
    totalVolumes = 0
                
2b) Loop over all the rows in the spreadsheet.
    
    Worksheets(yearValue).Activate
    For k = 2 To RowCount
        
3a) Increase volume for current ticker 

    If Cells(k, 1).Value = tickerIndex Then
    totalVolumes = totalVolumes + Cells(k, 8).Value
    End If
        
3b) Check if the current row is the first row with the selected tickerIndex.
        	
    If Cells(k, 1).Value = tickerIndex And Cells(k - 1, 1).Value <> tickerIndex Then
    tickerStartingPrices = Cells(k, 6).Value
    End If

3c) check if the current row is the last row with the selected ticker, If the next row’s ticker doesn’t match, increase the tickerIndex. If  Then,'set ending price, End If
            
    If Cells(k, 1).Value = tickerIndex And Cells(k + 1, 1).Value <> tickerIndex Then
    tickerEndingPrices = Cells(k, 6).Value
    End If

3d) Increase the tickerIndex.
      
      Next k
     
4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickerIndex
    Cells(4 + i, 2).Value = totalVolumes
    Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1

Next i

*Output and Code Performance*



<insert images>

### **Summary**

Refactoring code has several *advantages*:

* Can take less time than writing the code from scratch
* Code will be better organized
* Code could run more efficiently

But also some *disadvantages*:

* Poorly written code without comments may be difficult to follow
* When using someone else’s code you may unknowingly download something malicious
* It will not change the output code that is already operable and therefore may not always be necessary 

Refactoring the code for this project did not change the output, however it did improve the speed of the analysis.
