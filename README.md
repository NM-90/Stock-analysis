# Stock-analysis
Stock Analysis/Excel VBA

## Overview of Project
    To help Steve analyze 2017 and 2108 stock data more effeciently.  
    
### Purpose
    Steve wants to find the total daily volume and yearly return for all stocks in 2017 and 2018 using cleaner and better code. I used Excel VBA code to 
    refactor an earlier code to run the data faster.  
    
    Of note, this projext the word "stock" is replacted with "ticker".  Several tickers(stocks) were compared to see which ones gave the best yearly 
    returns(volume).  
    
## Results

### Analysis

    '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays

    Dim tickerVolumes(12) As Long

    Dim tickerStartingPrices(12) As Single

    Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.

    For i = 0 To 11

        tickerVolumes(i) = 0
    
        tickerStartingPrices(i) = 0
    
        tickerEndingPrices(i) = 0
    
    Next i

    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount

        '3a) Increase volume for current ticker
    
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
     '3b) Check if the current row is the first row with the selected tickerIndex.
    
     'If  Then
    
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
         tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
     End If
    
        '3c) check if the current row is the last row with the selected ticker
    
        'If the next row's ticker doesn't match, increase the tickerIndex.
   
        'If  Then
    
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
     
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
     End If

        '3d Increase the tickerIndex.
        
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         
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
    
After refactoring, these were the new run times:

<img width="229" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/111452227/200416257-63d0d9ff-a42d-4c89-852a-dddaa362b7aa.png">
<img width="218" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/111452227/200416286-b952e721-4410-4283-9a25-819139256abb.png">

### 

### Challenges and Difficulties Encountered

    Refactoring the code in general:
    One of the challenges of this project was learning how to loop the data.  Without properly looping the data, the output only provided the results of 
    of the first ticker.  It was imperative to include (i) in the loop. In turn, knowing when to close the loops, in this case there were two, "i" and   
    "j" and where to indent to give the the correct results. One also needs to inspect the data to understand what types of information there is in 
    order to create the correct arrays.  The tickerVolume was a Long number while the tickerStarting and Ending prices were Single numbers.  In general, 
    there difficutly was learning how to open the correct VBA macro while learning how to use VBA. Entering the code in Microsoft Excel Objects instead 
    of the Module sheets caused confusion.     
    
    Original and refactored VBA script:
    While refactoring the code provided a more efficient code, it did have some challenges.  First, the code had to be refined in order to produce both   
    2017 and 2018 using one macro instead of creating one for each year.  For reference, the initial code only gave the 2018 output. Another difference 
    in the refactored code was increasing of the ticker index, "tickerIndex = tickerIndex +1". And lastly, the tickerVolumes array was added including  
    the total number of the string (12) to the refactored code.  
    

