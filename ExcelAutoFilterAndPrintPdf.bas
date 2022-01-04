Sub ExcelAutoFilterAndPrintPdf()
'declare var's
Dim data As Worksheet
Dim list As Worksheet
Dim region As String
Dim count As Long
Dim i As Long

'get to be printed sheet

Set data = ThisWorkbook.Sheets(1)

'get second sheet were values of filter are present

Set list = ThisWorkbook.Sheets(2)

'count numnber of regions
'activate sheet of filter values

list.Activate

'count the total values to be filtered
'count = WorksheetFunction.CountA(Range("A1", Range("A1").End(xlDown)))
count = ActiveSheet.Cells(Rows.count, "A").End(xlUp).Row

'Display the filters count

MsgBox "Total Filter Values found is - " & count

'activate the main sheet to be printed

data.Activate

'start printing pdfs

For i = 1 To count
        
        'updating the region name and address
        region = list.Cells(i, 1).Text
        data.Cells(2, 1) = region
        
        'filter by current region
        Range("A4").AutoFilter field:=15, Criteria1:=region
        
        'save pdf
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=DirectoryLocation & _
         region & "_" & Format(Date, "dd-mmm-yyyy")

Next i

ActiveSheet.AutoFilterMode = False

End Sub

