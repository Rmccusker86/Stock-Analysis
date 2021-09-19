# Stock-Analysis

Purpose
The purpose of this project was to take code written in Microsoft Excel VBA and refractor it to show certain stock date for different years. This was done to determine whether or not the stocks had potential to invest in. During the module the code was written out, but the goal for this assignment was to re-write it to increase the efficiency.

The Results
Analysis
To start, I copied the code that was required to create the input box, chart headers, ticker arrays, and to activate the correct worksheet. Following the given instructions, I was able to layout and code the correct structure for refactoring. The results are shown below.

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
    
    Attached are the pictures showing the runtimes of the refactored code
    
    ![VBA_Challenge_2017](https://user-images.githubusercontent.com/88358771/133942443-fcc13117-71e6-4397-91ec-4615ba24be99.png)
    ![VBA_Challenge_2018](https://user-images.githubusercontent.com/88358771/133942450-e07c23a2-3d56-46c7-8d2a-6c8946858628.png)
    
    
    Pros and Cons of Refactoring code
    
    Firstly, refactoring code helps make the code cleaner and easier to understand. Other advantages include using less memory on the computer, as well as faster and more efficient debugging. Ultimately the greatest pro to a refactor of code is it would allow others to view the product be able to read and understand it, with it being more straightforward and concise. 
    For the cons side of a refactor, I would say that usually comes down to time alloted to go through line by line and spend the time cleaning up the code. Most job and real life scenarios either come with a time frame, or is expected to be donw as quickly as possible. As well as during the refactoring of the code, errors could be made which could ultimately break the Sub resulting in downtime for anyone who needs to use it.
    
    
    Conclusion on refactoring code
    
    With refactoring the code I was able to decrease the run time per year from .60 second down to .15 and .16 respectively. So spending the time to go in and clean the code resulted in cutting the Sub run time down to a quarter of its original. Overall, I would say that the advantages of refactoring a code to increase speed/efficiency of it far outweighs any disadvantages or risks associated with it. 



