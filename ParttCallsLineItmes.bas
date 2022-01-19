Sub PartCallslineItems()
'Formats the CSV file from Vcare site into a ready made working format
'Delete not required columns
Dim a As Long, w As Long, vDELCOLs As Variant, vCOLNDX As Variant
'Array
vDELCOLs = Array("SR Processing Status", "Row Id", "Order Description", "SA Type", "Deallocate Reason", "Row Id", "Returnable", "Part #", "SR Open Date", "SR Close Date", "Spare Invoice #", "Spare Invoice Date", "Spare Invoice Status", "Parent Invoice #", "Rejected By", "SF Age", "Challan Age", "Line Item Creation Date", "Capacity", "Created by Division", "Product Group", "H Status", "Type", "SAP Contract #", "L Status", "Contract Type", "Agreement", "Address", "Cancel Reason", "Customer Comments", "Remarks", "Escalation", "Severity", "VIP", "Mobile Update", "Defect Part #", "Defect Part Name", "Defect Return Status", "Manager", "Gas Charge Done Flag", "Part Required Flag", "Part Replaced Flag", "House #", "Building", "Road", "State", "Closure Code", "Purchased From", "Purchased From Free", "Last Modified By", "RT", "DT", "Attend time", "Appointment Date", "Serial# Source", "Split Serial# Source", "Serial Source Updated", "Split Serial Source Updated", "NPS Score", "Email Add", "Purchase Date")
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
    xStr_SR = "SR Number"
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
    
'Find Service_Agent_Code # column
    Dim xRg_Service_Agent_Code As Range
    Dim xRgUni_Service_Agent_Code As Range
    Dim xAddress_Service_Agent_Code As String
    Dim xStr_Service_Agent_Code As String
    On Error Resume Next
    xStr_Service_Agent_Code = "Franchisee Code"
    Set xRg_Service_Agent_Code = ActiveSheet.UsedRange.Find(xStr_Service_Agent_Code, , xlValues, xlWhole, , , True)
    If Not xRg_Service_Agent_Code Is Nothing Then
        xAddress_Service_Agent_Code = xRg_Service_Agent_Code.Address
        Do
            Set xRg_Service_Agent_Code = ActiveSheet.UsedRange.FindNext(xRg_Service_Agent_Code)
            If xRgUni_Service_Agent_Code Is Nothing Then
                Set xRgUni_Service_Agent_Code = xRg_Service_Agent_Code
            Else
                Set xRgUni_Service_Agent_Code = Application.Union(xRgUni_Service_Agent_Code, xRg_Service_Agent_Code)
            End If
        Loop While (Not xRg_Service_Agent_Code Is Nothing) And (xRg_Service_Agent_Code.Address <> xAddress_Service_Agent_Code)
    End If
    xRgUni_Service_Agent_Code.EntireColumn.Activate
     With xRgUni_Service_Agent_Code.EntireColumn
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
'Find Service_Agent_Code # column End
    
        
'Find Call_Type # column
    Dim xRg_Call_Type As Range
    Dim xRgUni_Call_Type As Range
    Dim xAddress_Call_Type As String
    Dim xStr_Call_Type As String
    On Error Resume Next
    xStr_Call_Type = "Call Type"
    Set xRg_Call_Type = ActiveSheet.UsedRange.Find(xStr_Call_Type, , xlValues, xlWhole, , , True)
    If Not xRg_Call_Type Is Nothing Then
        xAddress_Call_Type = xRg_Call_Type.Address
        Do
            Set xRg_Call_Type = ActiveSheet.UsedRange.FindNext(xRg_Call_Type)
            If xRgUni_Call_Type Is Nothing Then
                Set xRgUni_Call_Type = xRg_Call_Type
            Else
                Set xRgUni_Call_Type = Application.Union(xRgUni_Call_Type, xRg_Call_Type)
            End If
        Loop While (Not xRg_Call_Type Is Nothing) And (xRg_Call_Type.Address <> xAddress_Call_Type)
    End If
    xRgUni_Call_Type.EntireColumn.Activate
     With xRgUni_Call_Type.EntireColumn
        .ColumnWidth = 12
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
'Find Call_Type # column END


'Find Account # column
    Dim xRg_Account As Range
    Dim xRgUni_Account As Range
    Dim xAddress_Account As String
    Dim xStr_Account As String
    On Error Resume Next
    xStr_Account = "Account"
    Set xRg_Account = ActiveSheet.UsedRange.Find(xStr_Account, , xlValues, xlWhole, , , True)
    If Not xRg_Account Is Nothing Then
        xAddress_Account = xRg_Account.Address
        Do
            Set xRg_Account = ActiveSheet.UsedRange.FindNext(xRg_Account)
            If xRgUni_Account Is Nothing Then
                Set xRgUni_Account = xRg_Account
            Else
                Set xRgUni_Account = Application.Union(xRgUni_Account, xRg_Account)
            End If
        Loop While (Not xRg_Account Is Nothing) And (xRg_Account.Address <> xAddress_Account)
    End If
    xRgUni_Account.EntireColumn.Activate
     With xRgUni_Account.EntireColumn
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
'Find Account # column END


'Find Sub_Status # column
    Dim xRg_Sub_Status As Range
    Dim xRgUni_Sub_Status As Range
    Dim xAddress_Sub_Status As String
    Dim xStr_Sub_Status As String
    On Error Resume Next
    xStr_Sub_Status = "SR Sub Status"
    Set xRg_Sub_Status = ActiveSheet.UsedRange.Find(xStr_Sub_Status, , xlValues, xlWhole, , , True)
    If Not xRg_Sub_Status Is Nothing Then
        xAddress_Sub_Status = xRg_Sub_Status.Address
        Do
            Set xRg_Sub_Status = ActiveSheet.UsedRange.FindNext(xRg_Sub_Status)
            If xRgUni_Sub_Status Is Nothing Then
                Set xRgUni_Sub_Status = xRg_Sub_Status
            Else
                Set xRgUni_Sub_Status = Application.Union(xRgUni_Sub_Status, xRg_Sub_Status)
            End If
        Loop While (Not xRg_Sub_Status Is Nothing) And (xRg_Sub_Status.Address <> xAddress_Sub_Status)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Sub_Status.EntireColumn.Activate
     With xRgUni_Sub_Status.EntireColumn
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
        .FormatConditions.Delete
    End With
    
'Color foormating with string
'Re-opened calls
    With xRgUni_Sub_Status.EntireColumn.FormatConditions.Add(xlTextString, TextOperator:=xlContains, String:="Re-Opened")
        With .Font
            .Bold = True
            .ColorIndex = 3
        End With
    End With

    With xRgUni_Sub_Status.EntireColumn.FormatConditions.Add(xlTextString, TextOperator:=xlContains, String:="Cancel Request Rejected")
        With .Font
            .Bold = True
            .ColorIndex = 3
        End With
    End With
'Find Sub Status END

'SAP Order

'Find SAP ORDER # column
    Dim xRg_Field_1 As Range
    Dim xRgUni_Field_1 As Range
    Dim xAddress_Field_1 As String
    Dim xStr_Field_1 As String
    On Error Resume Next
    xStr_Field_1 = "SAP Order #"
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
'Find SAP Order # column END


'Find Order Sub Type # column
    Dim xRg_Field_2 As Range
    Dim xRgUni_Field_2 As Range
    Dim xAddress_Field_2 As String
    Dim xStr_Field_2 As String
    On Error Resume Next
    xStr_Field_2 = "Order Sub Type"
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
'Find Order Sub Type # column END

'Find SR Status # column
    Dim xRg_Field_3 As Range
    Dim xRgUni_Field_3 As Range
    Dim xAddress_Field_3 As String
    Dim xStr_Field_3 As String
    On Error Resume Next
    xStr_Field_3 = "SR Status"
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
        .ColumnWidth = 9
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
'Find SR Status # column END


'Find SAP Order Type# column
    Dim xRg_Field_4 As Range
    Dim xRgUni_Field_4 As Range
    Dim xAddress_Field_4 As String
    Dim xStr_Field_4 As String
    On Error Resume Next
    xStr_Field_4 = "SAP Order Type"
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
        .ColumnWidth = 9
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
'Find SAP Order Typecolumn END


'Find Order Number column
    Dim xRg_Field_5 As Range
    Dim xRgUni_Field_5 As Range
    Dim xAddress_Field_5 As String
    Dim xStr_Field_5 As String
    On Error Resume Next
    xStr_Field_5 = "Order Number"
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
'Find Order Number column END

'Find Service_Agent
    Dim xRg_Service_Agent As Range
    Dim xRgUni_Service_Agent As Range
    Dim xAddress_Service_Agent As String
    Dim xStr_Service_Agent As String
    On Error Resume Next
    xStr_Service_Agent = "Franchisee"
    Set xRg_Service_Agent = ActiveSheet.UsedRange.Find(xStr_Service_Agent, , xlValues, xlWhole, , , True)
    If Not xRg_Service_Agent Is Nothing Then
        xAddress_Service_Agent = xRg_Service_Agent.Address
        Do
            Set xRg_Service_Agent = ActiveSheet.UsedRange.FindNext(xRg_Service_Agent)
            If xRgUni_Service_Agent Is Nothing Then
                Set xRgUni_Service_Agent = xRg_Service_Agent
            Else
                Set xRgUni_Service_Agent = Application.Union(xRgUni_Service_Agent, xRg_Service_Agent)
            End If
        Loop While (Not xRg_Service_Agent Is Nothing) And (xRg_Service_Agent.Address <> xAddress_Service_Agent)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Service_Agent.EntireColumn.Activate
     With xRgUni_Service_Agent.EntireColumn
        .ColumnWidth = 25
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
        


'Find Order Date column
    Dim xRg_Field_6 As Range
    Dim xRgUni_Field_6 As Range
    Dim xAddress_Field_6 As String
    Dim xStr_Field_6 As String
    On Error Resume Next
    xStr_Field_6 = "Order Date"
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
'Find Order Date column END

        
'Find Shipment # column
    Dim xRg_Field_7 As Range
    Dim xRgUni_Field_7 As Range
    Dim xAddress_Field_7 As String
    Dim xStr_Field_7 As String
    On Error Resume Next
    xStr_Field_7 = "Shipment #"
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
'Find Shipment # column END


'Find Order Reason column
    Dim xRg_Field_8 As Range
    Dim xRgUni_Field_8 As Range
    Dim xAddress_Field_8 As String
    Dim xStr_Field_8 As String
    On Error Resume Next
    xStr_Field_8 = "Order Reason"
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
'Find Order Reason column END

'Find Product column
    Dim xRg_Field_9 As Range
    Dim xRgUni_Field_9 As Range
    Dim xAddress_Field_9 As String
    Dim xStr_Field_9 As String
    On Error Resume Next
    xStr_Field_9 = "Product"
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
'Find Product column END

'Find Quantity Requested column
    Dim xRg_Field_10 As Range
    Dim xRgUni_Field_10 As Range
    Dim xAddress_Field_10 As String
    Dim xStr_Field_10 As String
    On Error Resume Next
    xStr_Field_10 = "Quantity Requested"
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
'Find Quantity Requested column END

'Find Branch column
    Dim xRg_Field_11 As Range
    Dim xRgUni_Field_11 As Range
    Dim xAddress_Field_11 As String
    Dim xStr_Field_11 As String
    On Error Resume Next
    xStr_Field_11 = "Branch"
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
        .ColumnWidth = 9
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


'Find SAP Submission Date column
    Dim xRg_Field_12 As Range
    Dim xRgUni_Field_12 As Range
    Dim xAddress_Field_12 As String
    Dim xStr_Field_12 As String
    On Error Resume Next
    xStr_Field_12 = "SAP Submission Date"
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
'Find SAP Submission Date column END
        
'Find ETA/ Delivery Remarks column
    Dim xRg_Field_13 As Range
    Dim xRgUni_Field_13 As Range
    Dim xAddress_Field_13 As String
    Dim xStr_Field_13 As String
    On Error Resume Next
    xStr_Field_13 = "ETA/ Delivery Remarks"
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
        .ColumnWidth = 16
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
'Find ETA/ Delivery Remarks column END
       
       

'Find Registered Phone column
    Dim xRg_Field_14 As Range
    Dim xRgUni_Field_14 As Range
    Dim xAddress_Field_14 As String
    Dim xStr_Field_14 As String
    On Error Resume Next
    xStr_Field_14 = "Registered Phone"
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
'Find Registered Phone column END

    
'Find age # column
    Dim xRg_age As Range
    Dim xRgUni_age As Range
    Dim xAddress_age As String
    Dim xStr_age As String
    On Error Resume Next
    xStr_age = "SR Age"
    Set xRg_age = ActiveSheet.UsedRange.Find(xStr_age, , xlValues, xlWhole, , , True)
    If Not xRg_age Is Nothing Then
        xAddress_age = xRg_age.Address
        Do
            Set xRg_age = ActiveSheet.UsedRange.FindNext(xRg_age)
            If xRgUni_age Is Nothing Then
                Set xRgUni_age = xRg_age
            Else
                Set xRgUni_age = Application.Union(xRgUni_age, xRg_age)
            End If
        Loop While (Not xRg_age Is Nothing) And (xRg_age.Address <> xAddress_age)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_age.EntireColumn.Activate
     With xRgUni_age.EntireColumn
        .ColumnWidth = 4.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .FormatConditions.Delete
    End With
'Color foormating
'Add first rule
        xRgUni_age.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
                Formula1:="=3", Formula2:=" "
        xRgUni_age.EntireColumn.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        xRgUni_age.EntireColumn.FormatConditions(1).Font.Color = RGB(255, 255, 255)
        xRgUni_age.EntireColumn.FormatConditions(1).Font.Bold = True
        'Add second rule
        xRgUni_age.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                Formula1:="=2"
        xRgUni_age.EntireColumn.FormatConditions(2).Interior.Color = RGB(255, 128, 0)
        'Add third rule
        'xRgUni_age.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
            '   Formula1:="=0"
        'xRgUni_age.EntireColumn.FormatConditions(3).Interior.Color = vbWhite
        'Add fourth rule
        xRgUni_age.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                Formula1:="=1"
        xRgUni_age.EntireColumn.FormatConditions(3).Interior.Color = RGB(225, 229, 204)
        ' Add fifth rule
        'xRgUni_age.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        '       Formula1:="=0"
        'xRgUni_age.EntireColumn.FormatConditions(3).Interior.Color = vbYellow

'End of SR age


'Find Serial
    Dim xRg_Serial As Range
    Dim xRgUni_Serial As Range
    Dim xAddress_Serial As String
    Dim xStr_Serial As String
    On Error Resume Next
    xStr_Serial = "SR Serial Number"
    Set xRg_Serial = ActiveSheet.UsedRange.Find(xStr_Serial, , xlValues, xlWhole, , , True)
    If Not xRg_Serial Is Nothing Then
        xAddress_Serial = xRg_Serial.Address
        Do
            Set xRg_Serial = ActiveSheet.UsedRange.FindNext(xRg_Serial)
            If xRgUni_Serial Is Nothing Then
                Set xRgUni_Serial = xRg_Serial
            Else
                Set xRgUni_Serial = Application.Union(xRgUni_Serial, xRg_Serial)
            End If
        Loop While (Not xRg_Serial Is Nothing) And (xRg_Serial.Address <> xAddress_Serial)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Serial.EntireColumn.Activate
     With xRgUni_Serial.EntireColumn
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
    
'End Serial

'Find ASM
    Dim xRg_ASM As Range
    Dim xRgUni_ASM As Range
    Dim xAddress_ASM As String
    Dim xStr_ASM As String
    On Error Resume Next
    xStr_ASM = "ASM"
    Set xRg_ASM = ActiveSheet.UsedRange.Find(xStr_ASM, , xlValues, xlWhole, , , True)
    If Not xRg_ASM Is Nothing Then
        xAddress_ASM = xRg_ASM.Address
        Do
            Set xRg_ASM = ActiveSheet.UsedRange.FindNext(xRg_ASM)
            If xRgUni_ASM Is Nothing Then
                Set xRgUni_ASM = xRg_ASM
            Else
                Set xRgUni_ASM = Application.Union(xRgUni_ASM, xRg_ASM)
            End If
        Loop While (Not xRg_ASM Is Nothing) And (xRg_ASM.Address <> xAddress_ASM)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_ASM.EntireColumn.Activate
     With xRgUni_ASM.EntireColumn
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
    
    
'End ASM


'Find SR Product Category column
    Dim xRg_Field_15 As Range
    Dim xRgUni_Field_15 As Range
    Dim xAddress_Field_15 As String
    Dim xStr_Field_15 As String
    On Error Resume Next
    xStr_Field_15 = "SR Product Category"
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
        .ColumnWidth = 13.5
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
'Find SR Product Category column END

'Find Brand Identifier column
    Dim xRg_Field_16 As Range
    Dim xRgUni_Field_16 As Range
    Dim xAddress_Field_16 As String
    Dim xStr_Field_16 As String
    On Error Resume Next
    xStr_Field_16 = "Brand Identifier"
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


'Find Order Type column
    Dim xRg_Field_17 As Range
    Dim xRgUni_Field_17 As Range
    Dim xAddress_Field_17 As String
    Dim xStr_Field_17 As String
    On Error Resume Next
    xStr_Field_17 = "Order Type"
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
        .ColumnWidth = 11.5
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
'Find Order Type column END

'Find Order Part Type column
    Dim xRg_Field_18 As Range
    Dim xRgUni_Field_18 As Range
    Dim xAddress_Field_18 As String
    Dim xStr_Field_18 As String
    On Error Resume Next
    xStr_Field_18 = "Order Part Type"
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
'Find Order Part Type column END





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
  xRgUni_SR.EntireColumn.Activate
  
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









