Attribute VB_Name = "WriteTickerOutput"
Sub WriteTickerOutput( _
        Ticker As String, _
        YearlyChange As Double, _
        PercentChange As Double, _
        AnnualVolume As Double)


'Created by Jeff Pinegar
'Created Sept 27, 2022
'Last revision Date: Sept 27, 2022
'Assignment: VBA Challenge
'
' ----------------------------------------
'Purpose:
'This macro is use to write the ticker result into the results table
'Analysis of the Stock market table

'------------------------------------------

    'Variable declarations
    '-------------------------------------
'    Dim ticker As String
'    Dim YearlyChange As Double
'    Dim PercentChange As Double
'    Dim AnnualVolume As Double
    Dim rowNum As Long
   
    
    
    
    'test values - this should be commented out when complete
    '--------------------------------------
'    Ticker = "META"
'    YearlyChange = 3100
'    PercentChange = 0.119
'    AnnualVolume = 459123701
'    Worksheets("A").Activate
    
    
    
    'Find first blank cell in Column A
    '------------------------------------
    Application.ScreenUpdating = False          'stop the screen from flashing
    datasheet = ActiveSheet.Name                'where is the data comming from
    Worksheets("Output").Activate               'Activate the worksheet named Output.



    'find the first empty Row
    rowNum = ActiveSheet.Cells.Find( _
                What:="*", _
                After:=Cells(1, 1), _
                LookIn:=xlFormulas, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).row
    rowNum = rowNum + 1

    
    'The Ticker is use to find the blank row, therefore is no Ticker is passed enter "Missing Ticker"
    If Ticker = "" Then Ticker = "Missing Ticker"
    
    'fill in the table
    '-------------------------------------
    Cells(rowNum, 1).Value = datasheet
    Cells(rowNum, 2).Value = Ticker
    Cells(rowNum, 3).Value = YearlyChange
    Cells(rowNum, 4).Value = PercentChange
    Cells(rowNum, 5).Value = AnnualVolume
    
    
    'Activate the sheet that was being processed
    '------------------------------------------------------
    Worksheets(datasheet).Activate
    

End Sub

