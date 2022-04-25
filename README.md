# stock-analysis

**Analysis**

The Purpose of this project was to reduce the code used to be able to produce the same output and also make the software work at a higher efficiency. This resulted in mainly refactoring the loop to one loop instead of a previously nested loop solution

**Challenges**

1. Prepare the file for the project by changing the name of the file from green-stocks.xlsm to VBA_Challenge.xlsm (the challenge I faced here was to open the problem file since it was a VBscript and my computer was not providing notepad as the default software to open the file)
2. First create a new variable called tickerIndex then go on to create array's for tickerVolumes, tickerStartingPrices, and tickerEndingPrices
3. Create a loop to initially use the rowcount to increase volume, and then see if the current row matches the first row or last row, and if nothing matches then icnrease the tickerIndex
4. Final loop to return the Ticker, Total Daily Volume and Return in Cells A4 to C15 respectively

**Code Solution**

    ' Create a ticker Index
    
    tickerIndex = 0
    

    'Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    'Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
   
    
    Next i

        
    'Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        'Increase volume for current ticker
     tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
     
        
        ' Check if the current row is the first row with the selected tickerIndex.
        
       If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
       
       tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
       
       End If
        
        ' Check if the current row is the last row with the selected ticker, if the next row’s ticker doesn’t match, increase the tickerIndex.
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            

            ' Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerIndex = tickerIndex + 1
            
            
            End If
    
    Next i
    
    'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
        
    Next i
    
   **Advantages**
   
   The code runs significantly faster than when I previously ran it with more nested loops, and also takes up less lines (thus making it less time consuming to code)
   
   **Disadvantages**
   
   The previous code was more easier to explain to beginners and non coding folks. If we employ this current solution and show it to someone without a prior coding background they may stumble and not understand the overall code.
   
    
