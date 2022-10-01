Attribute VB_Name = "TickerTotal"
Sub TickerTotal(i As Integer)
'Created by Jeff Pinegar
'Created Sept 27, 2022
'Last revision Date: Sept 27, 2022
'Assignment: VBA Challenge
'
' ----------------------------------------
'Purpose:
'This macro is use to tabulate the total for a single ticker
'Analysis of the Stock market table

'------------------------------------------

    'Variable declarations
    '-------------------------------------
    Dim TabTicker As String             'This is the ticker we are tubulating
    Dim curTicker As String             'Ticker of the current row
    Dim Volume As Double                'Running total of volume for the year
    Dim startYearOpen As Double         'Opening price at start of year
    Dim deltaYear As Double             'Change in price first day open to last day close
    Dim deltaPercent As Double          'percent change in price
    Dim Sht As String                   'the name of the data worksheet we are processing
    
    
    
    Worksheets(i).Activate
    Sht = ActiveSheet.Name
    rowNum = 2              'data sheet start at row 2
    
    Do Until IsEmpty(Cells(rowNum, 1).Value)
            
            
            'Initialize the values for this pass
            '----------------------------------------
            Volume = 0                                  'reset the volume
            TabTicker = Cells(rowNum, 1).Value          'record the starting ticker
            curTicker = TabTicker                       'Current Ticker = Start ticker
            startYearOpen = Cells(rowNum, 3).Value      'record price at the begining of the year
            
            
            
            'sum the volume unitl the ticker changes
            '------------------------------------------
            Do Until curTicker <> TabTicker
                Volume = Volume + Cells(rowNum, 7).Value
                rowNum = rowNum + 1
                curTicker = Cells(rowNum, 1).Value
            Loop
        
        
            'Caculate final values
            '---------------------------------------------
            rowNum = rowNum - 1                             'return to the row with the final value
            deltaYear = Cells(rowNum, 6) - startYearOpen    'calculate the price change
            deltaPercent = deltaYear / startYearOpen        'calculate the percent change
            
            
            'Write the values to the ouput table
            '----------------------------------------------
            Call WriteTickerOutput.WriteTickerOutput(TabTicker, deltaYear, deltaPercent, Volume)
            
            Worksheets(i).Activate
            rowNum = rowNum + 1                         'Step down to the next potential ticker
            TabTicker = Cells(rowNum, 1).Value          'next ticker to tabulate
            
    Loop
            
End Sub
