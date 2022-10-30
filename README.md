# -stock-analysis-
Green_stocks
“Stock Analysis”

Overview of Project

In this project we help Steve analyze a stock DAQO for his parents, since they are going to be his first clients. He wanted to make sure that he would be able to provide information to analyze not just DAQO but many more green stocks so that his parents don’t spend their money in one stock, but that they can’t have more stock to pick from. With this VBA with one click on a button he would be able to gather all the information that was built in this VBA.

At first, he wanted us to build a excel VBA to read a dozen stock but then he wanted for u to refactor the project to read more than a dozen stocks. So, we have to refactor because our format was built for a dozen stock, it still works for more stocks, but it might take longer to processed. That is why h made us refactor the format to read stock for a few years behind.

Compressing of charts

2018
 

2017
 

The VBA  module
'1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
        
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(i) = 0
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
           tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
             tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
              tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
              
       

            '3d Increase the tickerIndex.
            
              tickerIndex = tickerIndex + 1
            
        End If
    
     Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
        
        Worksheets("AllStocksAnalysis").Activate
        
     Cells(4 + i, 1).Value = tickers(i)
     Cells(4 + i, 2).Value = tickerVolumes(i)
     Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i

Here we can use this code that ca be use to provide the information need to analysis for more then just dozen stock. In the chart we can see the information that Steve was asking for. With this refactor you can use it to run analysis for any year 2017 or 2018.

Advantage

Some of the advantage that is given with the Refactor is that it’s not time consuming it runs the code at a faster pace then the one made before. It’s a little clearer to so that anyone that is reading the code can fully understand it.

Disadvantage

	When I was refactoring this code it was really time consuming because I had to be very clear with the details. 
![image](https://user-images.githubusercontent.com/115046550/198884735-924aaf00-8d51-40f4-bf8e-ae2b34670d53.png)
