# Refactoring VBA Code

![Picture of Stock Market](https://mediacloud.kiplinger.com/image/private/s--x2_BoIgn--/v1604352227/Investing/stock-market-today-110220.jpg)

## **Overview of Project**

###  *Purpose* 

What information can be gathered quickly to make a sound decision around a financial investment? That is the true purpose of this analysis. By reviewing 2017 and 2018 stock market ticker data, the ability to determine which stocks would be a worthy investment based on daily performance become more apparent.

Additionally, there was a more technical added layer to the overarching goal of this project which was to fully understand the processes and skills required to refactor Visual Basic for Applications code within Microsoft Excel. The process of refactoring is critical when analyzing large sets of data on a regular basis. Through refactoring, data sets can be analyzed more efficiently, individuals reading and reviewing code are able to do so thoroughly through high quality formatting, and debugging occurs to ensure effectiveness of the code that is run.

## **Results**

Through processes of refactoring and debugging, the following code was produced to decrease the processing time of the analysis. 

### *Refactored Code* ###

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
    RowCount = Cells(Rows.Count, "A").End(xlUp).row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    Worksheets(yearValue).Activate
    For i = 0 To 11
        tickerVolumes(12) = 0
        tickerStartingPrices(12) = 0
        tickerEndingPrices(12) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
            '3d Increase the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
        'End If
        
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
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

After refactoring the code, the timer illustrated the improvement in run time for the overall macro. This process was conducted for both 2017 and 2018.

### *2017 Time After Refactoring* ###

<img width="259" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/87885677/133898485-579e4d44-8dd6-4654-bef7-9183428333db.png">

### *2018 Time After Refactoring* ###

<img width="255" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/87885677/133898513-6ea25b08-4d0a-41cb-832d-f9998d68295e.png">

## **Summary**
Which stocks make the most sense to invest in based on daily performance? Through a deeper analysis, there were two true standouts in 2018 - ENPH (81.9% return) and RUN (84.0% return). Although it could be misleading in 2017 that many stocks had incredible success, it is a safe decision to narrow in on the stocks identified above. Both experienced year after year of returns while the other stocks may be deemed more volative based on their 2018 return tumble after a lucrative 2017.

An added focus of this assignment was to evaluate the effectiveness of refactoring. The best illustration of this would be the time differential between the code prior to refactoring compared to after refactoring. For the 2017 data set, prior to refactoring the run time was approximately 0.292 seconds compared to 0.093 seconds in the refactored code. Additionally, for the 2018 data set, prior to refactoring the run time was also approximately 0.292 seconds compared to 0.089 seconds in the refactored code. Overall, the refactored code performed approximately 3 to 4 times faster than the original code. As noted earlier, the power of refactoring is not only in the quality of which the code is written and the readaibility, but also in the efficiency to generate high quality answers in a timely manner.

