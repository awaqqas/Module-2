# Module-2
Module 2 asisgnment with 2 deliverables
**Deliverable 1- Refactor VBA code and measure performance**
Purpose: Refactoring a stock analysis code to determine the type of stocks are worth investing into from the years 2017 and 2018. Code was refactored to ensure it ran faster while porviding accurate 'return' infomation on each stock. 
Dataset: Lists of 12 different stocks was provided in 2 separate sheets denoted by year "2017" and "2018".
Analysis:
Prior to refactoring the code, the code was copied from 'green_Stocks' file where first few lines of the code acted as the basis for the code:
startTime and endTime were intialized at the beginning to capture the time it took for the code to run on both '2017' and '2018' datasheets. 
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
  The code was refactored in the following manner:
	
  ''1a) Create a ticker Index
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
    
        
    ''2b) Looped over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increased volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Checked if the current row is the first row with the selected tickerIndex.
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) checked if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
            
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If


            '3d Increased the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    '4) Looped through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
       Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

**PROS and CONS of refactoring the code**
Rfeactoring the code made it more user-friendly to read and understand for the users. It also improves debugging, increases processing speed and enhances programming speed. While, the advantages of refactoring the code are present, it also requires developers with enough skills to be able to refactor a code for improved efficiency. In fact, in situations, where systems are integrating and scripts are running through millions of data-points, it may get very cumbersome to re-factor a code. Overall, involves immense amount of time and resources. 

Advantages of refactoring stock analysis code:
As mentioned earlier, one of the biggest advantages of refactoring a code is faster script running times which was apparent in our sotcks analysis.
The script run time reduces with each run as apparent between '2017' and '2018' run times. 

<img width="1437" alt="Screen Shot 2021-09-27 at 11 41 06 AM" src="https://user-images.githubusercontent.com/90429568/134940766-c93a78b9-cefd-4636-b356-5d90d92da686.png">

    
    <img width="1437" alt="Screen Shot 2021-09-27 at 11 42 27 AM" src="https://user-images.githubusercontent.com/90429568/134940987-56e2f3d4-c17b-46a5-a876-f846d47c64c0.png">

Additionally, the script had explanatory statements for each line of the code, enabling the reader to understand on how the code works (denoted by the green statements in the script). 
<img width="1437" alt="Screen Shot 2021-09-27 at 11 43 47 AM" src="https://user-images.githubusercontent.com/90429568/134941205-e7da767a-2f94-43e1-8d8b-8e4cd2d2e617.png">


