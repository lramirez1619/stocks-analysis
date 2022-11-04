# stocks-analysis
Learning VBA
# Module 2 Challenge

## Overview of Project
A good friend Steve and his family would like to invest on green energy, specifically, DAQO New Energy Corporation, however like any smart investor they would like additional data. They would like to be able to look at other potential investment and be able to run the same report more efficiently for other stock market over the last few years.

### Purpose
The main purpose of the challenge is to re-write or refactor the original report for DAQO to reflect other potential stock options while decreasing the runtime it takes for the report to run.

## Results

### Refactor: Module2_VBA_Script

    '1a) Create a ticker Index
     tickerindex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
          tickerVolumes(i) = 0
    
    Next i
     
    '2b) Loop over all the rows in the spreadsheet.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To RowCount
            
    '3a) Increase volume for current ticker
    tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, "H").Value
        
    '3b) Check if the current row is the first row with the selected tickerIndex.
    If Cells(i - 1, 1).Value <> tickers(tickerindex) Then
      
            tickerStartingPrices(tickerindex) = Cells(i, "F").Value
     
    End If
        
    '3c) check if the current row is the last row with the selected ticker
    'If the next rowâ€™s ticker doesn ot match, increase the tickerIndex.
     If Cells(i + 1, 1).Value <> tickers(tickerindex) Then

            tickerEndingPrices(tickerindex) = Cells(i, "F").Value
      
     '3d Increase the tickerIndex.
            tickerindex = tickerindex + 1
         
      End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, "A").Value = tickers(i)
        Cells(4 + i, "B").Value = tickerVolumes(i)
        Cells(4 + i, "C").Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

        
    Next i
    
### Original timer
![Original Run](https://github.com/lramirez1619/stocks-analysis/blob/33985fe08dad7407c8b5165e53953cced0eafd20/Resources/VBA_Challenge2.5_2018.png)

### Refactor timer
![Refactor Run](https://github.com/lramirez1619/stocks-analysis/blob/ff4b365911639b747a7dd4e793de4d5372043a8b/Resources/VBA_Challenge.25_2018.png)

### Analysis
In 2018, the market seems to be healthier in comparison to 2017 for green energy stocks. Based on the data, Steve and his family should consider investing on ENPH and RUN. For 2017, both ENPH and RUN did well in the market just like the other green energy stocks, however in 2018 ENPH experienced a 81.9% results while RUN experienced an 84% results. Both stocks did well however, the rest of the green energy stocks did not fare well and experienced a negative result in 2018. Both stocks did well historically, however a continual mornitoring would be wise as Steve and his family as they continue to invest in green energy.

## Summary

### Advantages of refactoring code
May not take as much time as building the codes from scratch. Allows you to look at the codes and re-organize or re-format in ways that is not only more user friendly but also requires less resource to run due to reduced code size. It is beneficial when onboarding a new user to an existing report. A good opportunity to add new features and functionalities as needed.
      
### Disadvantages of refactoring code
Refactoring may be a disadvantage when new bugs are introduced while in pursuit of reducing code sizes. It can also be a disadvatage if it causes to miss deadline since its an additional time investment. In some cases its possible that an original file is so poorly written that it may be more efficient to write a new set of codes than to try to decipher and re-write. 

### How do these pros and cons apply to refactoring the original VBA script?
In refactoring the original script, time reduced from 2.5 seconds to 0.26 seconds for 2018 data which is 89.6% time reduction. It is user friendly for future users.  The task did require time to execute.



