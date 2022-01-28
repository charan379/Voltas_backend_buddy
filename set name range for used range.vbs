Sub createNamedRange()
 
    'Source: https://powerspreadsheets.com/
    'For further information: https://powerspreadsheets.com/vba-create-named-range/
 
    'declare object variables to hold references to worksheet containing cell range, and cell range itself
    Dim myWorksheet As Worksheet
    Dim myNamedRange As range
 
    'declare variable to hold defined name
    Dim myRangeName As String
 
    'identify worksheet containing cell range, and cell range itself
  Set myWorksheet = ThisWorkbook.Worksheets("sheet name")
 
  'use below for active work sheet
 ' Set myWorksheet = ThisWorkbook.ActiveSheet
  
  
    Set myNamedRange = myWorksheet.UsedRange
 
    'specify defined name
  myRangeName = "Range Name"
 
    'create named range with workbook scope. Defined name and cell range are as specified
    ThisWorkbook.Names.Add Name:=myRangeName, RefersTo:=myNamedRange
    
End Sub


'as a funtion with args
Public Function CrNamedRange(Sname As String) As String
 
    'declare object variables to hold references to worksheet containing cell range, and cell range itself
    Dim myWorksheet As Worksheet
    Dim myNamedRange As Range
 
    'declare variable to hold defined name
    Dim myRangeName As String
 
    'identify worksheet containing cell range, and cell range itself
  Set myWorksheet = ThisWorkbook.Worksheets(Sname)
 
  'use below for active work sheet
 ' Set myWorksheet = ThisWorkbook.ActiveSheet
  
  
    Set myNamedRange = myWorksheet.UsedRange
 
    'specify defined name
  myRangeName = "callsRange"
 
    'create named range with workbook scope. Defined name and cell range are as specified
    ThisWorkbook.Names.Add Name:="callsRange", RefersTo:=myNamedRange
    
    
    'Returns Range name
    CrNamedRange = myRangeName

    
End Function
