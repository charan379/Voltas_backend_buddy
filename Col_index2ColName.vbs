Public Function Number2Letter(Col_index As Long) As String
'PURPOSE: Convert a given number into it's corresponding Letter Reference

Dim ColumnNumber As Long
Dim ColumnLetter As String

'Input Column Number
  ColumnNumber = Col_index

'Convert To Column Letter
  ColumnLetter = Split(Cells(1, ColumnNumber).Address, "$")(1)
  
'Display Result
  'MsgBox "Column " & ColumnNumber & " = Column " & ColumnLetter
  Number2Letter = ColumnLetter
  
End Function
