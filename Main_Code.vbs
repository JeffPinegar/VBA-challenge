Attribute VB_Name = "Main_Code"
Sub Main_code()
    Dim intShtTotal As Integer
    Dim intCurShtNum As Integer
    Dim NumOfRows As Long
    Dim i As Integer
    

    'create the sheet where the results table will be created
    Call CreateOutputSheet          'Output sheet is created as sheet #1

    'how many sheets of data are there?
    intShtTotal = Sheets.Count
    
    
    For i = 2 To intShtTotal      'It is 2 because we don't want to process the output sheet
        'find the number of rows in this table
        Worksheets(i).Activate
        NumOfRows = ActiveSheet.Cells.Find(What:="*", After:=Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).row
        
        'Sort the sheets
        Set Rng = Range(Cells(1, 2), Cells(NumOfRows, 1))
        Rng.Sort Key1:=Range("A1"), _
                 Order1:=xlAscending, _
                 Key2:=Range("B1"), _
                 Order1:=xlAscending, _
                 Header:=xlYes
               
        'Tabulate the sheets
        Call TickerTotal.TickerTotal(i)
        
        
        'Stop before going to next sheet - opportunity to kill program
        'MsgBox ("Sheet " & i & " Done")
        
        
    Next i
    
    'Stop before going to summary - opportunity to kill program
    'MsgBox ("Ready to Summarize Results")
    
    'Generate the summary result for each year
    Call GrtSummary
    Call OutputHeadings
    
    Range("A1").Select  'locate sheet at cell A1

'All done !!!
    
End Sub

Sub GrtSummary()
    
    Dim Tkr As String
    Dim GrtPerInc As Double
    Dim GrtPerDec As Double
    Dim GrtVol As Double
    Dim GrtPerIncTick As String
    Dim GrtPerDecTick As String
    Dim GrtVolTick As String
    Dim GrtOutputRow As Long
    Dim WkSht As String
    Dim row As Long                 'the current row
    Dim lastRowOfOutput As Long     'the last row of the output table

    'Switch to the output worksheet
    Worksheets("Output").Activate
    
    
    'find last row of the output table
    lastRowOfOutput = ActiveSheet.Cells.Find( _
                What:="*", _
                After:=Cells(1, 1), _
                LookIn:=xlFormulas, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).row
    
    
    'Sort the table the output table
    Set Rng = Range(Cells(1, 2), Cells(lastRowOfOutput, 5))
    Rng.Sort Key1:=Range("A1"), _
             Order1:=xlAscending, _
             Key2:=Range("B1"), _
             Order1:=xlAscending, _
             Header:=xlYes
    
    
    'Create heading in row 1 for the summary output
    Cells(1, 8).Value = "Year"         'WkSht cells(1,8)
    Cells(1, 9).Value = "Description            "   'Description cells(1,9)
    Cells(1, 9).ColumnWidth = 25
    Cells(1, 10).Value = "Ticker"       'Ticker cells(1,10)
    Cells(1, 11).Value = "Value"        'value cells(1,11)
    Cells(1, 11).ColumnWidth = 15
    
    
    
    'Make the summary output area into a table
    'Create Table named "GrtResultsTable"
    '----------------------------------------------
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$H$1:$K$1"), , xlYes).Name = "GrtResultsTable"
    ActiveSheet.ListObjects("GrtResultsTable").TableStyle = "TableStyleMedium11"
    
    
    
    'start at the top of the ouput table row 2
    'record the starting WkSht name - this is likely a year like 2018
    WkSht = Cells(2, 1).Value
    
    'record the starting row for the summary table set the first row for the first summary output
    GrtOutputRow = 2
    row = 2

    
    'Do-while loop next cell is not blank
    Do                                                          '*** Loop 1 ****  Loop 1 ****** Loop 1 **** Loop 1 ******
    
        'set initial conditions
        GrtPerInc = 0
        GrtPerDec = 0
        GrtVol = 0
        GrtPerIncTkr = 0
        GrtPerDecTkr = 0
        GrtVolTkr = 0

            'Do until WkSht changes
            Do Until WkSht <> Cells(row + 1, 1).Value           '*** Loop 2 ****  Loop 2 ****** Loop 2 **** Loop 2 ******
            
                'Percent change > GrtPerInc then save GrtPerInc and GrtPerIncTick
                If Cells(row, 4).Value > GrtPerInc Then
                    GrtPerInc = Cells(row, 4).Value
                    GrtPerIncTkr = Cells(row, 2).Value
                End If

                'Percent change < GrtPerDec then save GrtPerDec and GrtPerdecTick
                If Cells(row, 4).Value < GrtPerDec Then
                    GrtPerDec = Cells(row, 4).Value
                    GrtPerDecTkr = Cells(row, 2).Value
                End If

                'Volume > GrtVol then save GrtVol and GrtVolTick
                 If Cells(row, 5).Value > GrtVol Then
                    GrtVol = Cells(row, 5).Value
                    GrtVolTkr = Cells(row, 2).Value
                End If
                row = row + 1
                
            Loop                                                '*** Loop 2 ****  Loop 2 ****** Loop 2 **** Loop 2 ******
            
            'summary Greatest Percent Increase for the current worksheet = wksht
            Cells(GrtOutputRow, 8).Value = WkSht
            Cells(GrtOutputRow, 8).HorizontalAlignment = xlLeft
            Cells(GrtOutputRow, 9).Value = "Greatest % Increase"
            Cells(GrtOutputRow, 10).Value = GrtPerIncTkr
            Cells(GrtOutputRow, 11).Value = GrtPerInc
            Cells(GrtOutputRow, 11).NumberFormat = "0.00%"
            GrtOutputRow = GrtOutputRow + 1

            'summary Greatest Percent Decrease for the current worksheet = wksht
            Cells(GrtOutputRow, 8).Value = WkSht
            Cells(GrtOutputRow, 8).HorizontalAlignment = xlLeft
            Cells(GrtOutputRow, 9).Value = "Greatest % Decrease"
            Cells(GrtOutputRow, 10).Value = GrtPerDecTkr
            Cells(GrtOutputRow, 11).Value = GrtPerDec
            Cells(GrtOutputRow, 11).NumberFormat = "0.00%"
            GrtOutputRow = GrtOutputRow + 1

            'summary Greatest volume for the current worksheet = wksht
            Cells(GrtOutputRow, 8).Value = WkSht
            Cells(GrtOutputRow, 8).HorizontalAlignment = xlLeft
            Cells(GrtOutputRow, 9).Value = "Greatest Total Volume"
            Cells(GrtOutputRow, 10).Value = GrtVolTkr
            Cells(GrtOutputRow, 11).Value = GrtVol
            Cells(GrtOutputRow, 11).NumberFormat = "#,##0"
            GrtOutputRow = GrtOutputRow + 1

             
            'go to next worksheet
            row = row + 1
            WkSht = Cells(row, 1).Value

    'Loop While the next worksheet name is not blank
    Loop While Not (IsEmpty(Cells(row, 1).Value))                     '*** Loop 1 ****  Loop 1 ****** Loop 1 **** Loop 1 ******

End Sub


Sub OutputHeadings()
'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Ticker Summary"
    With Selection.Font
        .Name = "Calibri"
        .Size = 22
        .Bold = True
    End With
    
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Yearly Summary"
    With Selection.Font
        .Name = "Calibri"
        .Size = 22
        .Bold = True
    End With
    Selection.Font.Bold = True
End Sub

