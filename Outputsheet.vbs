Attribute VB_Name = "Outputsheet"
Sub CreateOutputSheet()
Attribute CreateOutputSheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
'Created by Jeff Pinegar
'Created Sept 26, 2022
'Last revision Date: Sept 26, 2022
'Assignment: VBA Challenge
'
' ----------------------------------------
'Purpose:
'This macro is use to create and format the worksheet used to report the
'Analysis of the Stock market table

'------------------------------------------

    'Dim ResultsTable As TableObject
'
    'Add a sheet named output as the first sheet.
    '----------------------------------------------
    Sheets.Add Before:=Sheets(1)
    Sheets(1).Name = "Output"
    
    'Add the headings
   '----------------------------------------------
    Cells(1, 1).Value = "Year"
    Cells(1, 2).Value = "Ticker"
    Cells(1, 3).Value = "Yearly Change"
    Cells(1, 4).Value = "Percent Change"
    Cells(1, 5).Value = "Total Stock Volume"
    
    'Create Table named "ResultsTable"
    '----------------------------------------------
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$E$1"), , xlYes).Name = "ResultsTable"
    ActiveSheet.ListObjects("ResultsTable").TableStyle = "TableStyleMedium9"
    
    'Make conditional format red with white text if "Yearly Change is less than 0
    '----------------------------------------------
    Range("C2").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Interior.ColorIndex = 3       'Red backgroung
    Selection.FormatConditions(1).Font.ColorIndex = 2           'White font
        
       
    'Make conditional format green if "Yearly Change is less greater than or equal to 0
    '----------------------------------------------
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Interior.ColorIndex = 35      'Green Background
    
    'Establish numeric formats for the columns
    '--------------------------------------------
    Range("A2").Select
    Selection.HorizontalAlignment = xlLeft
    Range("c2").Select
    Selection.NumberFormat = "$#,##0.00"
    Range("d2").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    Range("E2").Select
    Selection.NumberFormat = "#,##0"
    
    
    'Postion for first Entry
    '----------------------------------------------
    Range("A2").Select
     
End Sub


