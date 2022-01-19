Sub DefctiveOrderDashboard()

'Formats the CSV file from Vcare site into a ready made working format
'Delete not required columns
Dim a As Long, w As Long, vDELCOLs As Variant, vCOLNDX As Variant
'Array
vDELCOLs = Array("SR Closed Date", "Godam Defect Serial #", "Spare Invoice Number", "Spare Invoice Status", "Spare Invoice Date", "Revised Manual Challan", "Order Line Item#", "SAP Order Type", "SAP Order Ref", "SR Unit Status", "Order Date", "SAP Submission Date", "Line Item Status", "Order Header Status", "Attachment", "Godam Docket", "SAP Inbound Doc Num", "SAP PO Num", "SAP PO Date", "Org part INV Num", "Org part INV Value", "Spares FPO Value", "HO Remarks", "Sr Purchase Date", "Ok Part Serial Num", "SF Type", "Claim Date", "Claim Reason", "Self Courier", "Quantity Requested", "Related Product", "Related Product Name", "Error Log", "Transport", "SR Call Type", "Submitted By", "Submitted Time", "Branch Code", "SR Status")
With ThisWorkbook
    For w = 1 To .Worksheets.count
        With Worksheets(w)
            For a = LBound(vDELCOLs) To UBound(vDELCOLs)
                vCOLNDX = Application.Match(vDELCOLs(a), .Rows(1), 0)
                If Not IsError(vCOLNDX) Then
                    .Columns(vCOLNDX).EntireColumn.Delete
                End If
            Next a
        End With
    Next w
End With

'Delete not required columns end

'Find SR # column
    Dim xRg_SR As Range
    Dim xRgUni_SR As Range
    Dim xAddress_SR As String
    Dim xStr_SR As String
    On Error Resume Next
    xStr_SR = "SR #"
    Set xRg_SR = ActiveSheet.UsedRange.Find(xStr_SR, , xlValues, xlWhole, , , True)
    If Not xRg_SR Is Nothing Then
        xAddress_SR = xRg_SR.Address
        Do
            Set xRg_SR = ActiveSheet.UsedRange.FindNext(xRg_SR)
            If xRgUni_SR Is Nothing Then
                Set xRgUni_SR = xRg_SR
            Else
                Set xRgUni_SR = Application.Union(xRgUni_SR, xRg_SR)
            End If
        Loop While (Not xRg_SR Is Nothing) And (xRg_SR.Address <> xAddress_SR)
    End If
    xRgUni_SR.EntireColumn.Activate
     With xRgUni_SR.EntireColumn
        .AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find SR # column End

'Find Brand Identifier column
    Dim xRg_Field_1 As Range
    Dim xRgUni_Field_1 As Range
    Dim xAddress_Field_1 As String
    Dim xStr_Field_1 As String
    On Error Resume Next
    xStr_Field_1 = "Brand Identifier"
    Set xRg_Field_1 = ActiveSheet.UsedRange.Find(xStr_Field_1, , xlValues, xlWhole, , , True)
    If Not xRg_Field_1 Is Nothing Then
        xAddress_Field_1 = xRg_Field_1.Address
        Do
            Set xRg_Field_1 = ActiveSheet.UsedRange.FindNext(xRg_Field_1)
            If xRgUni_Field_1 Is Nothing Then
                Set xRgUni_Field_1 = xRg_Field_1
            Else
                Set xRgUni_Field_1 = Application.Union(xRgUni_Field_1, xRg_Field_1)
            End If
        Loop While (Not xRg_Field_1 Is Nothing) And (xRg_Field_1.Address <> xAddress_Field_1)
    End If
    xRgUni_Field_1.EntireColumn.Activate
     With xRgUni_Field_1.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Brand Identifier column END


'Find Branch column
    Dim xRg_Field_2 As Range
    Dim xRgUni_Field_2 As Range
    Dim xAddress_Field_2 As String
    Dim xStr_Field_2 As String
    On Error Resume Next
    xStr_Field_2 = "Branch"
    Set xRg_Field_2 = ActiveSheet.UsedRange.Find(xStr_Field_2, , xlValues, xlWhole, , , True)
    If Not xRg_Field_2 Is Nothing Then
        xAddress_Field_2 = xRg_Field_2.Address
        Do
            Set xRg_Field_2 = ActiveSheet.UsedRange.FindNext(xRg_Field_2)
            If xRgUni_Field_2 Is Nothing Then
                Set xRgUni_Field_2 = xRg_Field_2
            Else
                Set xRgUni_Field_2 = Application.Union(xRgUni_Field_2, xRg_Field_2)
            End If
        Loop While (Not xRg_Field_2 Is Nothing) And (xRg_Field_2.Address <> xAddress_Field_2)
    End If
    xRgUni_Field_2.EntireColumn.Activate
     With xRgUni_Field_2.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Branch column END

'Find Franchisee Name column
    Dim xRg_Field_3 As Range
    Dim xRgUni_Field_3 As Range
    Dim xAddress_Field_3 As String
    Dim xStr_Field_3 As String
    On Error Resume Next
    xStr_Field_3 = "Franchisee Name"
    Set xRg_Field_3 = ActiveSheet.UsedRange.Find(xStr_Field_3, , xlValues, xlWhole, , , True)
    If Not xRg_Field_3 Is Nothing Then
        xAddress_Field_3 = xRg_Field_3.Address
        Do
            Set xRg_Field_3 = ActiveSheet.UsedRange.FindNext(xRg_Field_3)
            If xRgUni_Field_3 Is Nothing Then
                Set xRgUni_Field_3 = xRg_Field_3
            Else
                Set xRgUni_Field_3 = Application.Union(xRgUni_Field_3, xRg_Field_3)
            End If
        Loop While (Not xRg_Field_3 Is Nothing) And (xRg_Field_3.Address <> xAddress_Field_3)
    End If
    xRgUni_Field_3.EntireColumn.Activate
     With xRgUni_Field_3.EntireColumn
        .ColumnWidth = 20
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Franchisee Name column END

'Find Franchisee Code column
    Dim xRg_Field_4 As Range
    Dim xRgUni_Field_4 As Range
    Dim xAddress_Field_4 As String
    Dim xStr_Field_4 As String
    On Error Resume Next
    xStr_Field_4 = "Franchisee Code"
    Set xRg_Field_4 = ActiveSheet.UsedRange.Find(xStr_Field_4, , xlValues, xlWhole, , , True)
    If Not xRg_Field_4 Is Nothing Then
        xAddress_Field_4 = xRg_Field_4.Address
        Do
            Set xRg_Field_4 = ActiveSheet.UsedRange.FindNext(xRg_Field_4)
            If xRgUni_Field_4 Is Nothing Then
                Set xRgUni_Field_4 = xRg_Field_4
            Else
                Set xRgUni_Field_4 = Application.Union(xRgUni_Field_4, xRg_Field_4)
            End If
        Loop While (Not xRg_Field_4 Is Nothing) And (xRg_Field_4.Address <> xAddress_Field_4)
    End If
    xRgUni_Field_4.EntireColumn.Activate
     With xRgUni_Field_4.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Franchisee Code column END


'Find SR Sub Status column
    Dim xRg_Field_5 As Range
    Dim xRgUni_Field_5 As Range
    Dim xAddress_Field_5 As String
    Dim xStr_Field_5 As String
    On Error Resume Next
    xStr_Field_5 = "SR Sub Status"
    Set xRg_Field_5 = ActiveSheet.UsedRange.Find(xStr_Field_5, , xlValues, xlWhole, , , True)
    If Not xRg_Field_5 Is Nothing Then
        xAddress_Field_5 = xRg_Field_5.Address
        Do
            Set xRg_Field_5 = ActiveSheet.UsedRange.FindNext(xRg_Field_5)
            If xRgUni_Field_5 Is Nothing Then
                Set xRgUni_Field_5 = xRg_Field_5
            Else
                Set xRgUni_Field_5 = Application.Union(xRgUni_Field_5, xRg_Field_5)
            End If
        Loop While (Not xRg_Field_5 Is Nothing) And (xRg_Field_5.Address <> xAddress_Field_5)
    End If
    xRgUni_Field_5.EntireColumn.Activate
     With xRgUni_Field_5.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find SR Sub Status column END

'Find SR Processing Status column
    Dim xRg_Field_6 As Range
    Dim xRgUni_Field_6 As Range
    Dim xAddress_Field_6 As String
    Dim xStr_Field_6 As String
    On Error Resume Next
    xStr_Field_6 = "SR Processing Status"
    Set xRg_Field_6 = ActiveSheet.UsedRange.Find(xStr_Field_6, , xlValues, xlWhole, , , True)
    If Not xRg_Field_6 Is Nothing Then
        xAddress_Field_6 = xRg_Field_6.Address
        Do
            Set xRg_Field_6 = ActiveSheet.UsedRange.FindNext(xRg_Field_6)
            If xRgUni_Field_6 Is Nothing Then
                Set xRgUni_Field_6 = xRg_Field_6
            Else
                Set xRgUni_Field_6 = Application.Union(xRgUni_Field_6, xRg_Field_6)
            End If
        Loop While (Not xRg_Field_6 Is Nothing) And (xRg_Field_6.Address <> xAddress_Field_6)
    End If
    xRgUni_Field_6.EntireColumn.Activate
     With xRgUni_Field_6.EntireColumn
        .ColumnWidth = 15
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find SR Processing Status column END


'Find SR Audit status column
    Dim xRg_Field_7 As Range
    Dim xRgUni_Field_7 As Range
    Dim xAddress_Field_7 As String
    Dim xStr_Field_7 As String
    On Error Resume Next
    xStr_Field_7 = "SR Audit status"
    Set xRg_Field_7 = ActiveSheet.UsedRange.Find(xStr_Field_7, , xlValues, xlWhole, , , True)
    If Not xRg_Field_7 Is Nothing Then
        xAddress_Field_7 = xRg_Field_7.Address
        Do
            Set xRg_Field_7 = ActiveSheet.UsedRange.FindNext(xRg_Field_7)
            If xRgUni_Field_7 Is Nothing Then
                Set xRgUni_Field_7 = xRg_Field_7
            Else
                Set xRgUni_Field_7 = Application.Union(xRgUni_Field_7, xRg_Field_7)
            End If
        Loop While (Not xRg_Field_7 Is Nothing) And (xRg_Field_7.Address <> xAddress_Field_7)
    End If
    xRgUni_Field_7.EntireColumn.Activate
     With xRgUni_Field_7.EntireColumn
        .ColumnWidth = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find SR Audit status column END

'Find SR Processed Date column
    Dim xRg_Field_8 As Range
    Dim xRgUni_Field_8 As Range
    Dim xAddress_Field_8 As String
    Dim xStr_Field_8 As String
    On Error Resume Next
    xStr_Field_8 = "SR Processed Date"
    Set xRg_Field_8 = ActiveSheet.UsedRange.Find(xStr_Field_8, , xlValues, xlWhole, , , True)
    If Not xRg_Field_8 Is Nothing Then
        xAddress_Field_8 = xRg_Field_8.Address
        Do
            Set xRg_Field_8 = ActiveSheet.UsedRange.FindNext(xRg_Field_8)
            If xRgUni_Field_8 Is Nothing Then
                Set xRgUni_Field_8 = xRg_Field_8
            Else
                Set xRgUni_Field_8 = Application.Union(xRgUni_Field_8, xRg_Field_8)
            End If
        Loop While (Not xRg_Field_8 Is Nothing) And (xRg_Field_8.Address <> xAddress_Field_8)
    End If
    xRgUni_Field_8.EntireColumn.Activate
     With xRgUni_Field_8.EntireColumn
     .ColumnWidth = 15
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .NumberFormat = "dd-mm-yyyy"
    End With
'Find SR Processed Date column END

'Find SR Account Name column
    Dim xRg_Field_9 As Range
    Dim xRgUni_Field_9 As Range
    Dim xAddress_Field_9 As String
    Dim xStr_Field_9 As String
    On Error Resume Next
    xStr_Field_9 = "SR Account Name"
    Set xRg_Field_9 = ActiveSheet.UsedRange.Find(xStr_Field_9, , xlValues, xlWhole, , , True)
    If Not xRg_Field_9 Is Nothing Then
        xAddress_Field_9 = xRg_Field_9.Address
        Do
            Set xRg_Field_9 = ActiveSheet.UsedRange.FindNext(xRg_Field_9)
            If xRgUni_Field_9 Is Nothing Then
                Set xRgUni_Field_9 = xRg_Field_9
            Else
                Set xRgUni_Field_9 = Application.Union(xRgUni_Field_9, xRg_Field_9)
            End If
        Loop While (Not xRg_Field_9 Is Nothing) And (xRg_Field_9.Address <> xAddress_Field_9)
    End If
    xRgUni_Field_9.EntireColumn.Activate
     With xRgUni_Field_9.EntireColumn
        .ColumnWidth = 14.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find SR Account Name column END


'Find SR Registered Number column
    Dim xRg_Field_10 As Range
    Dim xRgUni_Field_10 As Range
    Dim xAddress_Field_10 As String
    Dim xStr_Field_10 As String
    On Error Resume Next
    xStr_Field_10 = "SR Registered Number"
    Set xRg_Field_10 = ActiveSheet.UsedRange.Find(xStr_Field_10, , xlValues, xlWhole, , , True)
    If Not xRg_Field_10 Is Nothing Then
        xAddress_Field_10 = xRg_Field_10.Address
        Do
            Set xRg_Field_10 = ActiveSheet.UsedRange.FindNext(xRg_Field_10)
            If xRgUni_Field_10 Is Nothing Then
                Set xRgUni_Field_10 = xRg_Field_10
            Else
                Set xRgUni_Field_10 = Application.Union(xRgUni_Field_10, xRg_Field_10)
            End If
        Loop While (Not xRg_Field_10 Is Nothing) And (xRg_Field_10.Address <> xAddress_Field_10)
    End If
    xRgUni_Field_10.EntireColumn.Activate
     With xRgUni_Field_10.EntireColumn
        .ColumnWidth = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find SR Registered Number column END



'Find Sales Order Number column
    Dim xRg_Field_11 As Range
    Dim xRgUni_Field_11 As Range
    Dim xAddress_Field_11 As String
    Dim xStr_Field_11 As String
    On Error Resume Next
    xStr_Field_11 = "Sales Order Number"
    Set xRg_Field_11 = ActiveSheet.UsedRange.Find(xStr_Field_11, , xlValues, xlWhole, , , True)
    If Not xRg_Field_11 Is Nothing Then
        xAddress_Field_11 = xRg_Field_11.Address
        Do
            Set xRg_Field_11 = ActiveSheet.UsedRange.FindNext(xRg_Field_11)
            If xRgUni_Field_11 Is Nothing Then
                Set xRgUni_Field_11 = xRg_Field_11
            Else
                Set xRgUni_Field_11 = Application.Union(xRgUni_Field_11, xRg_Field_11)
            End If
        Loop While (Not xRg_Field_11 Is Nothing) And (xRg_Field_11.Address <> xAddress_Field_11)
    End If
    xRgUni_Field_11.EntireColumn.Activate
     With xRgUni_Field_11.EntireColumn
        .ColumnWidth = 14.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Sales Order Number column END


'Find Order Sub Type column
    Dim xRg_Field_12 As Range
    Dim xRgUni_Field_12 As Range
    Dim xAddress_Field_12 As String
    Dim xStr_Field_12 As String
    On Error Resume Next
    xStr_Field_12 = "Order Sub Type"
    Set xRg_Field_12 = ActiveSheet.UsedRange.Find(xStr_Field_12, , xlValues, xlWhole, , , True)
    If Not xRg_Field_12 Is Nothing Then
        xAddress_Field_12 = xRg_Field_12.Address
        Do
            Set xRg_Field_12 = ActiveSheet.UsedRange.FindNext(xRg_Field_12)
            If xRgUni_Field_12 Is Nothing Then
                Set xRgUni_Field_12 = xRg_Field_12
            Else
                Set xRgUni_Field_12 = Application.Union(xRgUni_Field_12, xRg_Field_12)
            End If
        Loop While (Not xRg_Field_12 Is Nothing) And (xRg_Field_12.Address <> xAddress_Field_12)
    End If
    xRgUni_Field_12.EntireColumn.Activate
     With xRgUni_Field_12.EntireColumn
        .ColumnWidth = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Order Sub Type column END


'Find SAP Order column
    Dim xRg_Field_13 As Range
    Dim xRgUni_Field_13 As Range
    Dim xAddress_Field_13 As String
    Dim xStr_Field_13 As String
    On Error Resume Next
    xStr_Field_13 = "SAP Order"
    Set xRg_Field_13 = ActiveSheet.UsedRange.Find(xStr_Field_13, , xlValues, xlWhole, , , True)
    If Not xRg_Field_13 Is Nothing Then
        xAddress_Field_13 = xRg_Field_13.Address
        Do
            Set xRg_Field_13 = ActiveSheet.UsedRange.FindNext(xRg_Field_13)
            If xRgUni_Field_13 Is Nothing Then
                Set xRgUni_Field_13 = xRg_Field_13
            Else
                Set xRgUni_Field_13 = Application.Union(xRgUni_Field_13, xRg_Field_13)
            End If
        Loop While (Not xRg_Field_13 Is Nothing) And (xRg_Field_13.Address <> xAddress_Field_13)
    End If
    xRgUni_Field_13.EntireColumn.Activate
     With xRgUni_Field_13.EntireColumn
        .ColumnWidth = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find SAP Order column END


'Find Order Reason column
    Dim xRg_Field_14 As Range
    Dim xRgUni_Field_14 As Range
    Dim xAddress_Field_14 As String
    Dim xStr_Field_14 As String
    On Error Resume Next
    xStr_Field_14 = "Order Reason"
    Set xRg_Field_14 = ActiveSheet.UsedRange.Find(xStr_Field_14, , xlValues, xlWhole, , , True)
    If Not xRg_Field_14 Is Nothing Then
        xAddress_Field_14 = xRg_Field_14.Address
        Do
            Set xRg_Field_14 = ActiveSheet.UsedRange.FindNext(xRg_Field_14)
            If xRgUni_Field_14 Is Nothing Then
                Set xRgUni_Field_14 = xRg_Field_14
            Else
                Set xRgUni_Field_14 = Application.Union(xRgUni_Field_14, xRg_Field_14)
            End If
        Loop While (Not xRg_Field_14 Is Nothing) And (xRg_Field_14.Address <> xAddress_Field_14)
    End If
    xRgUni_Field_14.EntireColumn.Activate
     With xRgUni_Field_14.EntireColumn
        .ColumnWidth = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Order Reason column END

'Find Manual Challan column
    Dim xRg_Field_15 As Range
    Dim xRgUni_Field_15 As Range
    Dim xAddress_Field_15 As String
    Dim xStr_Field_15 As String
    On Error Resume Next
    xStr_Field_15 = "Manual Challan"
    Set xRg_Field_15 = ActiveSheet.UsedRange.Find(xStr_Field_15, , xlValues, xlWhole, , , True)
    If Not xRg_Field_15 Is Nothing Then
        xAddress_Field_15 = xRg_Field_15.Address
        Do
            Set xRg_Field_15 = ActiveSheet.UsedRange.FindNext(xRg_Field_15)
            If xRgUni_Field_15 Is Nothing Then
                Set xRgUni_Field_15 = xRg_Field_15
            Else
                Set xRgUni_Field_15 = Application.Union(xRgUni_Field_15, xRg_Field_15)
            End If
        Loop While (Not xRg_Field_15 Is Nothing) And (xRg_Field_15.Address <> xAddress_Field_15)
    End If
    xRgUni_Field_15.EntireColumn.Activate
     With xRgUni_Field_15.EntireColumn
        .ColumnWidth = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Manual Challan column END


'Find Docket #/ LR No column
    Dim xRg_Field_16 As Range
    Dim xRgUni_Field_16 As Range
    Dim xAddress_Field_16 As String
    Dim xStr_Field_16 As String
    On Error Resume Next
    xStr_Field_16 = "Docket #/ LR No"
    Set xRg_Field_16 = ActiveSheet.UsedRange.Find(xStr_Field_16, , xlValues, xlWhole, , , True)
    If Not xRg_Field_16 Is Nothing Then
        xAddress_Field_16 = xRg_Field_16.Address
        Do
            Set xRg_Field_16 = ActiveSheet.UsedRange.FindNext(xRg_Field_16)
            If xRgUni_Field_16 Is Nothing Then
                Set xRgUni_Field_16 = xRg_Field_16
            Else
                Set xRgUni_Field_16 = Application.Union(xRgUni_Field_16, xRg_Field_16)
            End If
        Loop While (Not xRg_Field_16 Is Nothing) And (xRg_Field_16.Address <> xAddress_Field_16)
    End If
    xRgUni_Field_16.EntireColumn.Activate
     With xRgUni_Field_16.EntireColumn
        .ColumnWidth = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Docket #/ LR No column END


'Find Docket Date column
    Dim xRg_Field_17 As Range
    Dim xRgUni_Field_17 As Range
    Dim xAddress_Field_17 As String
    Dim xStr_Field_17 As String
    On Error Resume Next
    xStr_Field_17 = "Docket Date"
    Set xRg_Field_17 = ActiveSheet.UsedRange.Find(xStr_Field_17, , xlValues, xlWhole, , , True)
    If Not xRg_Field_17 Is Nothing Then
        xAddress_Field_17 = xRg_Field_17.Address
        Do
            Set xRg_Field_17 = ActiveSheet.UsedRange.FindNext(xRg_Field_17)
            If xRgUni_Field_17 Is Nothing Then
                Set xRgUni_Field_17 = xRg_Field_17
            Else
                Set xRgUni_Field_17 = Application.Union(xRgUni_Field_17, xRg_Field_17)
            End If
        Loop While (Not xRg_Field_17 Is Nothing) And (xRg_Field_17.Address <> xAddress_Field_17)
    End If
    xRgUni_Field_17.EntireColumn.Activate
     With xRgUni_Field_17.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .NumberFormat = "dd-mm-yyyy"
    End With
'Find Docket Date column END


'Find WH Remarks column
    Dim xRg_Field_18 As Range
    Dim xRgUni_Field_18 As Range
    Dim xAddress_Field_18 As String
    Dim xStr_Field_18 As String
    On Error Resume Next
    xStr_Field_18 = "WH Remarks"
    Set xRg_Field_18 = ActiveSheet.UsedRange.Find(xStr_Field_18, , xlValues, xlWhole, , , True)
    If Not xRg_Field_18 Is Nothing Then
        xAddress_Field_18 = xRg_Field_18.Address
        Do
            Set xRg_Field_18 = ActiveSheet.UsedRange.FindNext(xRg_Field_18)
            If xRgUni_Field_18 Is Nothing Then
                Set xRgUni_Field_18 = xRg_Field_18
            Else
                Set xRgUni_Field_18 = Application.Union(xRgUni_Field_18, xRg_Field_18)
            End If
        Loop While (Not xRg_Field_18 Is Nothing) And (xRg_Field_18.Address <> xAddress_Field_18)
    End If
    xRgUni_Field_18.EntireColumn.Activate
     With xRgUni_Field_18.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find WH Remarks column END


'Find Defective Serial # column
    Dim xRg_Field_19 As Range
    Dim xRgUni_Field_19 As Range
    Dim xAddress_Field_19 As String
    Dim xStr_Field_19 As String
    On Error Resume Next
    xStr_Field_19 = "Defective Serial #"
    Set xRg_Field_19 = ActiveSheet.UsedRange.Find(xStr_Field_19, , xlValues, xlWhole, , , True)
    If Not xRg_Field_19 Is Nothing Then
        xAddress_Field_19 = xRg_Field_19.Address
        Do
            Set xRg_Field_19 = ActiveSheet.UsedRange.FindNext(xRg_Field_19)
            If xRgUni_Field_19 Is Nothing Then
                Set xRgUni_Field_19 = xRg_Field_19
            Else
                Set xRgUni_Field_19 = Application.Union(xRgUni_Field_19, xRg_Field_19)
            End If
        Loop While (Not xRg_Field_19 Is Nothing) And (xRg_Field_19.Address <> xAddress_Field_19)
    End If
    xRgUni_Field_19.EntireColumn.Activate
     With xRgUni_Field_19.EntireColumn
        .ColumnWidth = 17
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Defective Serial # column END

'Find Defect Part Status column
    Dim xRg_Field_20 As Range
    Dim xRgUni_Field_20 As Range
    Dim xAddress_Field_20 As String
    Dim xStr_Field_20 As String
    On Error Resume Next
    xStr_Field_20 = "Defect Part Status"
    Set xRg_Field_20 = ActiveSheet.UsedRange.Find(xStr_Field_20, , xlValues, xlWhole, , , True)
    If Not xRg_Field_20 Is Nothing Then
        xAddress_Field_20 = xRg_Field_20.Address
        Do
            Set xRg_Field_20 = ActiveSheet.UsedRange.FindNext(xRg_Field_20)
            If xRgUni_Field_20 Is Nothing Then
                Set xRgUni_Field_20 = xRg_Field_20
            Else
                Set xRgUni_Field_20 = Application.Union(xRgUni_Field_20, xRg_Field_20)
            End If
        Loop While (Not xRg_Field_20 Is Nothing) And (xRg_Field_20.Address <> xAddress_Field_20)
    End If
    xRgUni_Field_20.EntireColumn.Activate
     With xRgUni_Field_20.EntireColumn
        .ColumnWidth = 15
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Defect Part Status column END


'Find Original Challan# column
    Dim xRg_Field_21 As Range
    Dim xRgUni_Field_21 As Range
    Dim xAddress_Field_21 As String
    Dim xStr_Field_21 As String
    On Error Resume Next
    xStr_Field_21 = "Original Challan#"
    Set xRg_Field_21 = ActiveSheet.UsedRange.Find(xStr_Field_21, , xlValues, xlWhole, , , True)
    If Not xRg_Field_21 Is Nothing Then
        xAddress_Field_21 = xRg_Field_21.Address
        Do
            Set xRg_Field_21 = ActiveSheet.UsedRange.FindNext(xRg_Field_21)
            If xRgUni_Field_21 Is Nothing Then
                Set xRgUni_Field_21 = xRg_Field_21
            Else
                Set xRgUni_Field_21 = Application.Union(xRgUni_Field_21, xRg_Field_21)
            End If
        Loop While (Not xRg_Field_21 Is Nothing) And (xRg_Field_21.Address <> xAddress_Field_21)
    End If
    xRgUni_Field_21.EntireColumn.Activate
     With xRgUni_Field_21.EntireColumn
        .ColumnWidth = 13
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Original Challan# column END


'Find Original Challan Date column
    Dim xRg_Field_22 As Range
    Dim xRgUni_Field_22 As Range
    Dim xAddress_Field_22 As String
    Dim xStr_Field_22 As String
    On Error Resume Next
    xStr_Field_22 = "Original Challan Date"
    Set xRg_Field_22 = ActiveSheet.UsedRange.Find(xStr_Field_22, , xlValues, xlWhole, , , True)
    If Not xRg_Field_22 Is Nothing Then
        xAddress_Field_22 = xRg_Field_22.Address
        Do
            Set xRg_Field_22 = ActiveSheet.UsedRange.FindNext(xRg_Field_22)
            If xRgUni_Field_22 Is Nothing Then
                Set xRgUni_Field_22 = xRg_Field_22
            Else
                Set xRgUni_Field_22 = Application.Union(xRgUni_Field_22, xRg_Field_22)
            End If
        Loop While (Not xRg_Field_22 Is Nothing) And (xRg_Field_22.Address <> xAddress_Field_22)
    End If
    xRgUni_Field_22.EntireColumn.Activate
     With xRgUni_Field_22.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .NumberFormat = "dd-mm-yyyy"
    End With
'Find Original Challan Date column END


'Find Rejected Date column
    Dim xRg_Field_23 As Range
    Dim xRgUni_Field_23 As Range
    Dim xAddress_Field_23 As String
    Dim xStr_Field_23 As String
    On Error Resume Next
    xStr_Field_23 = "Rejected Date"
    Set xRg_Field_23 = ActiveSheet.UsedRange.Find(xStr_Field_23, , xlValues, xlWhole, , , True)
    If Not xRg_Field_23 Is Nothing Then
        xAddress_Field_23 = xRg_Field_23.Address
        Do
            Set xRg_Field_23 = ActiveSheet.UsedRange.FindNext(xRg_Field_23)
            If xRgUni_Field_23 Is Nothing Then
                Set xRgUni_Field_23 = xRg_Field_23
            Else
                Set xRgUni_Field_23 = Application.Union(xRgUni_Field_23, xRg_Field_23)
            End If
        Loop While (Not xRg_Field_23 Is Nothing) And (xRg_Field_23.Address <> xAddress_Field_23)
    End If
    xRgUni_Field_23.EntireColumn.Activate
     With xRgUni_Field_23.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .NumberFormat = "dd-mm-yyyy"
    End With
'Find Rejected Date column END


'Find Approved Date column
    Dim xRg_Field_24 As Range
    Dim xRgUni_Field_24 As Range
    Dim xAddress_Field_24 As String
    Dim xStr_Field_24 As String
    On Error Resume Next
    xStr_Field_24 = "Approved Date"
    Set xRg_Field_24 = ActiveSheet.UsedRange.Find(xStr_Field_24, , xlValues, xlWhole, , , True)
    If Not xRg_Field_24 Is Nothing Then
        xAddress_Field_24 = xRg_Field_24.Address
        Do
            Set xRg_Field_24 = ActiveSheet.UsedRange.FindNext(xRg_Field_24)
            If xRgUni_Field_24 Is Nothing Then
                Set xRgUni_Field_24 = xRg_Field_24
            Else
                Set xRgUni_Field_24 = Application.Union(xRgUni_Field_24, xRg_Field_24)
            End If
        Loop While (Not xRg_Field_24 Is Nothing) And (xRg_Field_24.Address <> xAddress_Field_24)
    End If
    xRgUni_Field_24.EntireColumn.Activate
     With xRgUni_Field_24.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .NumberFormat = "dd-mm-yyyy"
    End With
'Find Approved Date column END


'Find Expired Date column
    Dim xRg_Field_25 As Range
    Dim xRgUni_Field_25 As Range
    Dim xAddress_Field_25 As String
    Dim xStr_Field_25 As String
    On Error Resume Next
    xStr_Field_25 = "Expired Date"
    Set xRg_Field_25 = ActiveSheet.UsedRange.Find(xStr_Field_25, , xlValues, xlWhole, , , True)
    If Not xRg_Field_25 Is Nothing Then
        xAddress_Field_25 = xRg_Field_25.Address
        Do
            Set xRg_Field_25 = ActiveSheet.UsedRange.FindNext(xRg_Field_25)
            If xRgUni_Field_25 Is Nothing Then
                Set xRgUni_Field_25 = xRg_Field_25
            Else
                Set xRgUni_Field_25 = Application.Union(xRgUni_Field_25, xRg_Field_25)
            End If
        Loop While (Not xRg_Field_25 Is Nothing) And (xRg_Field_25.Address <> xAddress_Field_25)
    End If
    xRgUni_Field_25.EntireColumn.Activate
     With xRgUni_Field_25.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .NumberFormat = "dd-mm-yyyy"
    End With
'Find Expired Date column END


'Find Returnable column
    Dim xRg_Field_26 As Range
    Dim xRgUni_Field_26 As Range
    Dim xAddress_Field_26 As String
    Dim xStr_Field_26 As String
    On Error Resume Next
    xStr_Field_26 = "Returnable"
    Set xRg_Field_26 = ActiveSheet.UsedRange.Find(xStr_Field_26, , xlValues, xlWhole, , , True)
    If Not xRg_Field_26 Is Nothing Then
        xAddress_Field_26 = xRg_Field_26.Address
        Do
            Set xRg_Field_26 = ActiveSheet.UsedRange.FindNext(xRg_Field_26)
            If xRgUni_Field_26 Is Nothing Then
                Set xRgUni_Field_26 = xRg_Field_26
            Else
                Set xRgUni_Field_26 = Application.Union(xRgUni_Field_26, xRg_Field_26)
            End If
        Loop While (Not xRg_Field_26 Is Nothing) And (xRg_Field_26.Address <> xAddress_Field_26)
    End If
    xRgUni_Field_26.EntireColumn.Activate
     With xRgUni_Field_26.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        '.NumberFormat = "dd-mm-yyyy"
    End With
'Find Returnable column END

'Find Revised Challan# column
    Dim xRg_Field_27 As Range
    Dim xRgUni_Field_27 As Range
    Dim xAddress_Field_27 As String
    Dim xStr_Field_27 As String
    On Error Resume Next
    xStr_Field_27 = "Revised Challan#"
    Set xRg_Field_27 = ActiveSheet.UsedRange.Find(xStr_Field_27, , xlValues, xlWhole, , , True)
    If Not xRg_Field_27 Is Nothing Then
        xAddress_Field_27 = xRg_Field_27.Address
        Do
            Set xRg_Field_27 = ActiveSheet.UsedRange.FindNext(xRg_Field_27)
            If xRgUni_Field_27 Is Nothing Then
                Set xRgUni_Field_27 = xRg_Field_27
            Else
                Set xRgUni_Field_27 = Application.Union(xRgUni_Field_27, xRg_Field_27)
            End If
        Loop While (Not xRg_Field_27 Is Nothing) And (xRg_Field_27.Address <> xAddress_Field_27)
    End If
    xRgUni_Field_27.EntireColumn.Activate
     With xRgUni_Field_27.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Revised Challan# column END


'Find Revised Challan Date column
    Dim xRg_Field_28 As Range
    Dim xRgUni_Field_28 As Range
    Dim xAddress_Field_28 As String
    Dim xStr_Field_28 As String
    On Error Resume Next
    xStr_Field_28 = "Revised Challan Date"
    Set xRg_Field_28 = ActiveSheet.UsedRange.Find(xStr_Field_28, , xlValues, xlWhole, , , True)
    If Not xRg_Field_28 Is Nothing Then
        xAddress_Field_28 = xRg_Field_28.Address
        Do
            Set xRg_Field_28 = ActiveSheet.UsedRange.FindNext(xRg_Field_28)
            If xRgUni_Field_28 Is Nothing Then
                Set xRgUni_Field_28 = xRg_Field_28
            Else
                Set xRgUni_Field_28 = Application.Union(xRgUni_Field_28, xRg_Field_28)
            End If
        Loop While (Not xRg_Field_28 Is Nothing) And (xRg_Field_28.Address <> xAddress_Field_28)
    End If
    xRgUni_Field_28.EntireColumn.Activate
     With xRgUni_Field_28.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .NumberFormat = "dd-mm-yyyy"
    End With
'Find Revised Challan Date column END


'Find Revised Docket# column
    Dim xRg_Field_29 As Range
    Dim xRgUni_Field_29 As Range
    Dim xAddress_Field_29 As String
    Dim xStr_Field_29 As String
    On Error Resume Next
    xStr_Field_29 = "Revised Docket#"
    Set xRg_Field_29 = ActiveSheet.UsedRange.Find(xStr_Field_29, , xlValues, xlWhole, , , True)
    If Not xRg_Field_29 Is Nothing Then
        xAddress_Field_29 = xRg_Field_29.Address
        Do
            Set xRg_Field_29 = ActiveSheet.UsedRange.FindNext(xRg_Field_29)
            If xRgUni_Field_29 Is Nothing Then
                Set xRgUni_Field_29 = xRg_Field_29
            Else
                Set xRgUni_Field_29 = Application.Union(xRgUni_Field_29, xRg_Field_29)
            End If
        Loop While (Not xRg_Field_29 Is Nothing) And (xRg_Field_29.Address <> xAddress_Field_29)
    End If
    xRgUni_Field_29.EntireColumn.Activate
     With xRgUni_Field_29.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        '.NumberFormat = "dd-mm-yyyy"
    End With
'Find Revised Docket# column END

'Find Revised Docket Date column
    Dim xRg_Field_30 As Range
    Dim xRgUni_Field_30 As Range
    Dim xAddress_Field_30 As String
    Dim xStr_Field_30 As String
    On Error Resume Next
    xStr_Field_30 = "Revised Docket Date"
    Set xRg_Field_30 = ActiveSheet.UsedRange.Find(xStr_Field_30, , xlValues, xlWhole, , , True)
    If Not xRg_Field_30 Is Nothing Then
        xAddress_Field_30 = xRg_Field_30.Address
        Do
            Set xRg_Field_30 = ActiveSheet.UsedRange.FindNext(xRg_Field_30)
            If xRgUni_Field_30 Is Nothing Then
                Set xRgUni_Field_30 = xRg_Field_30
            Else
                Set xRgUni_Field_30 = Application.Union(xRgUni_Field_30, xRg_Field_30)
            End If
        Loop While (Not xRg_Field_30 Is Nothing) And (xRg_Field_30.Address <> xAddress_Field_30)
    End If
    xRgUni_Field_30.EntireColumn.Activate
     With xRgUni_Field_30.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       .NumberFormat = "dd-mm-yyyy"
    End With
'Find Revised Docket Date column END

'Find Godam GRN Date column
    Dim xRg_Field_31 As Range
    Dim xRgUni_Field_31 As Range
    Dim xAddress_Field_31 As String
    Dim xStr_Field_31 As String
    On Error Resume Next
    xStr_Field_31 = "Godam GRN Date"
    Set xRg_Field_31 = ActiveSheet.UsedRange.Find(xStr_Field_31, , xlValues, xlWhole, , , True)
    If Not xRg_Field_31 Is Nothing Then
        xAddress_Field_31 = xRg_Field_31.Address
        Do
            Set xRg_Field_31 = ActiveSheet.UsedRange.FindNext(xRg_Field_31)
            If xRgUni_Field_31 Is Nothing Then
                Set xRgUni_Field_31 = xRg_Field_31
            Else
                Set xRgUni_Field_31 = Application.Union(xRgUni_Field_31, xRg_Field_31)
            End If
        Loop While (Not xRg_Field_31 Is Nothing) And (xRg_Field_31.Address <> xAddress_Field_31)
    End If
    xRgUni_Field_31.EntireColumn.Activate
     With xRgUni_Field_31.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       .NumberFormat = "dd-mm-yyyy"
    End With
'Find Godam GRN Date column END

'Find Godam Rejection Reason column
    Dim xRg_Field_32 As Range
    Dim xRgUni_Field_32 As Range
    Dim xAddress_Field_32 As String
    Dim xStr_Field_32 As String
    On Error Resume Next
    xStr_Field_32 = "Godam Rejection Reason"
    Set xRg_Field_32 = ActiveSheet.UsedRange.Find(xStr_Field_32, , xlValues, xlWhole, , , True)
    If Not xRg_Field_32 Is Nothing Then
        xAddress_Field_32 = xRg_Field_32.Address
        Do
            Set xRg_Field_32 = ActiveSheet.UsedRange.FindNext(xRg_Field_32)
            If xRgUni_Field_32 Is Nothing Then
                Set xRgUni_Field_32 = xRg_Field_32
            Else
                Set xRgUni_Field_32 = Application.Union(xRgUni_Field_32, xRg_Field_32)
            End If
        Loop While (Not xRg_Field_32 Is Nothing) And (xRg_Field_32.Address <> xAddress_Field_32)
    End If
    xRgUni_Field_32.EntireColumn.Activate
     With xRgUni_Field_32.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       '.NumberFormat = "dd-mm-yyyy"
    End With
'Find Godam Rejection Reason column END

'Find Godam Remarks column
    Dim xRg_Field_33 As Range
    Dim xRgUni_Field_33 As Range
    Dim xAddress_Field_33 As String
    Dim xStr_Field_33 As String
    On Error Resume Next
    xStr_Field_33 = "Godam Remarks"
    Set xRg_Field_33 = ActiveSheet.UsedRange.Find(xStr_Field_33, , xlValues, xlWhole, , , True)
    If Not xRg_Field_33 Is Nothing Then
        xAddress_Field_33 = xRg_Field_33.Address
        Do
            Set xRg_Field_33 = ActiveSheet.UsedRange.FindNext(xRg_Field_33)
            If xRgUni_Field_33 Is Nothing Then
                Set xRgUni_Field_33 = xRg_Field_33
            Else
                Set xRgUni_Field_33 = Application.Union(xRgUni_Field_33, xRg_Field_33)
            End If
        Loop While (Not xRg_Field_33 Is Nothing) And (xRg_Field_33.Address <> xAddress_Field_33)
    End If
    xRgUni_Field_33.EntireColumn.Activate
     With xRgUni_Field_33.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       '.NumberFormat = "dd-mm-yyyy"
    End With
'Find Godam Remarks column END

'Find Serialized Flag column
    Dim xRg_Field_34 As Range
    Dim xRgUni_Field_34 As Range
    Dim xAddress_Field_34 As String
    Dim xStr_Field_34 As String
    On Error Resume Next
    xStr_Field_34 = "Serialized Flag"
    Set xRg_Field_34 = ActiveSheet.UsedRange.Find(xStr_Field_34, , xlValues, xlWhole, , , True)
    If Not xRg_Field_34 Is Nothing Then
        xAddress_Field_34 = xRg_Field_34.Address
        Do
            Set xRg_Field_34 = ActiveSheet.UsedRange.FindNext(xRg_Field_34)
            If xRgUni_Field_34 Is Nothing Then
                Set xRgUni_Field_34 = xRg_Field_34
            Else
                Set xRgUni_Field_34 = Application.Union(xRgUni_Field_34, xRg_Field_34)
            End If
        Loop While (Not xRg_Field_34 Is Nothing) And (xRg_Field_34.Address <> xAddress_Field_34)
    End If
    xRgUni_Field_34.EntireColumn.Activate
     With xRgUni_Field_34.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       '.NumberFormat = "dd-mm-yyyy"
    End With
'Find Serialized Flag column END

'Find SR Product Category column
    Dim xRg_Field_35 As Range
    Dim xRgUni_Field_35 As Range
    Dim xAddress_Field_35 As String
    Dim xStr_Field_35 As String
    On Error Resume Next
    xStr_Field_35 = "SR Product Category"
    Set xRg_Field_35 = ActiveSheet.UsedRange.Find(xStr_Field_35, , xlValues, xlWhole, , , True)
    If Not xRg_Field_35 Is Nothing Then
        xAddress_Field_35 = xRg_Field_35.Address
        Do
            Set xRg_Field_35 = ActiveSheet.UsedRange.FindNext(xRg_Field_35)
            If xRgUni_Field_35 Is Nothing Then
                Set xRgUni_Field_35 = xRg_Field_35
            Else
                Set xRgUni_Field_35 = Application.Union(xRgUni_Field_35, xRg_Field_35)
            End If
        Loop While (Not xRg_Field_35 Is Nothing) And (xRg_Field_35.Address <> xAddress_Field_35)
    End If
    xRgUni_Field_35.EntireColumn.Activate
     With xRgUni_Field_35.EntireColumn
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       '.NumberFormat = "dd-mm-yyyy"
    End With
'Find SR Product Category column END

'Find Challan Age column
    Dim xRg_Field_36 As Range
    Dim xRgUni_Field_36 As Range
    Dim xAddress_Field_36 As String
    Dim xStr_Field_36 As String
    On Error Resume Next
    xStr_Field_36 = "Challan Age"
    Set xRg_Field_36 = ActiveSheet.UsedRange.Find(xStr_Field_36, , xlValues, xlWhole, , , True)
    If Not xRg_Field_36 Is Nothing Then
        xAddress_Field_36 = xRg_Field_36.Address
        Do
            Set xRg_Field_36 = ActiveSheet.UsedRange.FindNext(xRg_Field_36)
            If xRgUni_Field_36 Is Nothing Then
                Set xRgUni_Field_36 = xRg_Field_36
            Else
                Set xRgUni_Field_36 = Application.Union(xRgUni_Field_36, xRg_Field_36)
            End If
        Loop While (Not xRg_Field_36 Is Nothing) And (xRg_Field_36.Address <> xAddress_Field_36)
    End If
    xRgUni_Field_36.EntireColumn.Activate
     With xRgUni_Field_36.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       '.NumberFormat = "dd-mm-yyyy"
    End With
'Find Challan Age column END


'Find Product# column
    Dim xRg_Field_37 As Range
    Dim xRgUni_Field_37 As Range
    Dim xAddress_Field_37 As String
    Dim xStr_Field_37 As String
    On Error Resume Next
    xStr_Field_37 = "Product#"
    Set xRg_Field_37 = ActiveSheet.UsedRange.Find(xStr_Field_37, , xlValues, xlWhole, , , True)
    If Not xRg_Field_37 Is Nothing Then
        xAddress_Field_37 = xRg_Field_37.Address
        Do
            Set xRg_Field_37 = ActiveSheet.UsedRange.FindNext(xRg_Field_37)
            If xRgUni_Field_37 Is Nothing Then
                Set xRgUni_Field_37 = xRg_Field_37
            Else
                Set xRgUni_Field_37 = Application.Union(xRgUni_Field_37, xRg_Field_37)
            End If
        Loop While (Not xRg_Field_37 Is Nothing) And (xRg_Field_37.Address <> xAddress_Field_37)
    End If
    xRgUni_Field_37.EntireColumn.Activate
     With xRgUni_Field_37.EntireColumn
        .ColumnWidth = 15
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       '.NumberFormat = "dd-mm-yyyy"
    End With
'Find Product# column END

'Find SF Age column
    Dim xRg_Field_38 As Range
    Dim xRgUni_Field_38 As Range
    Dim xAddress_Field_38 As String
    Dim xStr_Field_38 As String
    On Error Resume Next
    xStr_Field_38 = "SF Age"
    Set xRg_Field_38 = ActiveSheet.UsedRange.Find(xStr_Field_38, , xlValues, xlWhole, , , True)
    If Not xRg_Field_38 Is Nothing Then
        xAddress_Field_38 = xRg_Field_38.Address
        Do
            Set xRg_Field_38 = ActiveSheet.UsedRange.FindNext(xRg_Field_38)
            If xRgUni_Field_38 Is Nothing Then
                Set xRgUni_Field_38 = xRg_Field_38
            Else
                Set xRgUni_Field_38 = Application.Union(xRgUni_Field_38, xRg_Field_38)
            End If
        Loop While (Not xRg_Field_38 Is Nothing) And (xRg_Field_38.Address <> xAddress_Field_38)
    End If
    xRgUni_Field_38.EntireColumn.Activate
     With xRgUni_Field_38.EntireColumn
        .AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       '.NumberFormat = "dd-mm-yyyy"
    End With
'Find SF Age column END

'Find Technician Name column
    Dim xRg_Field_39 As Range
    Dim xRgUni_Field_39 As Range
    Dim xAddress_Field_39 As String
    Dim xStr_Field_39 As String
    On Error Resume Next
    xStr_Field_39 = "Technician Name"
    Set xRg_Field_39 = ActiveSheet.UsedRange.Find(xStr_Field_39, , xlValues, xlWhole, , , True)
    If Not xRg_Field_39 Is Nothing Then
        xAddress_Field_39 = xRg_Field_39.Address
        Do
            Set xRg_Field_39 = ActiveSheet.UsedRange.FindNext(xRg_Field_39)
            If xRgUni_Field_39 Is Nothing Then
                Set xRgUni_Field_39 = xRg_Field_39
            Else
                Set xRgUni_Field_39 = Application.Union(xRgUni_Field_39, xRg_Field_39)
            End If
        Loop While (Not xRg_Field_39 Is Nothing) And (xRg_Field_39.Address <> xAddress_Field_39)
    End If
    xRgUni_Field_39.EntireColumn.Activate
     With xRgUni_Field_39.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       '.NumberFormat = "dd-mm-yyyy"
    End With
'Find Technician Name column END


'Find Base Price column
    Dim xRg_Field_40 As Range
    Dim xRgUni_Field_40 As Range
    Dim xAddress_Field_40 As String
    Dim xStr_Field_40 As String
    On Error Resume Next
    xStr_Field_40 = "Base Price"
    Set xRg_Field_40 = ActiveSheet.UsedRange.Find(xStr_Field_40, , xlValues, xlWhole, , , True)
    If Not xRg_Field_40 Is Nothing Then
        xAddress_Field_40 = xRg_Field_40.Address
        Do
            Set xRg_Field_40 = ActiveSheet.UsedRange.FindNext(xRg_Field_40)
            If xRgUni_Field_40 Is Nothing Then
                Set xRgUni_Field_40 = xRg_Field_40
            Else
                Set xRgUni_Field_40 = Application.Union(xRgUni_Field_40, xRg_Field_40)
            End If
        Loop While (Not xRg_Field_40 Is Nothing) And (xRg_Field_40.Address <> xAddress_Field_40)
    End If
    xRgUni_Field_40.EntireColumn.Activate
     With xRgUni_Field_40.EntireColumn
        .AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
       '.NumberFormat = "dd-mm-yyyy"
    End With
'Find Base Price column END




'------------------------------------------------  Top Row Color and page setup------

'Add Borders
 With ActiveSheet.UsedRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
 End With
  
'Color For Top row
With ActiveSheet.Range("A1", Cells(1, Columns.count).End(xlToRight)).SpecialCells(xlCellTypeConstants)
  .Interior.ColorIndex = 6
  .Font.Bold = True
End With
 xRgUni_SR.Activate
  
  ' Page layout set up
  With ActiveSheet.PageSetup
     .Orientation = xlLandscape
     .PaperSize = xlPaperA4
     '.Zoom = 80
     .Zoom = False
     '.FitToPagesTall = True
     '.FitToPagesWide = True
     .LeftMargin = Application.InchesToPoints(0.35)
    .RightMargin = Application.InchesToPoints(0.35)
    .TopMargin = Application.InchesToPoints(0.35)
    .BottomMargin = Application.InchesToPoints(0.35)
     .HeaderMargin = Application.InchesToPoints(0.35)
     .FooterMargin = Application.InchesToPoints(0.35)
     '.DisplayPageBreaks = False
     
End With

Dim ws As Worksheet
 
    For Each ws In ThisWorkbook.Worksheets
        ws.DisplayPageBreaks = False
        ws.ResetAllPageBreaks
 
    Next ws
    
    
    
    
    
End Sub
