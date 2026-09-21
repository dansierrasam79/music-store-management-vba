Attribute VB_Name = "modSales"
Option Explicit

Public Sub AddSale(ByVal sale As clsSale)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Set ws = ThisWorkbook.Worksheets("Sales")
    Set tbl = ws.ListObjects("tblsls")
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = sale.SaleID
        .Cells(1, 2).Value = sale.RecordID
        .Cells(1, 3).Value = sale.AlbumTitle
        .Cells(1, 4).Value = sale.QuantitySold
        .Cells(1, 5).Value = sale.SalePrice
        .Cells(1, 6).Value = sale.SaleDate
        .Cells(1, 7).Value = sale.CustomerID
    End With
End Sub

Public Function GetSaleByID(ByVal SaleID As Long) As clsSale
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim sale As clsSale
    Set ws = ThisWorkbook.Worksheets("Sales")
    Set tbl = ws.ListObjects("tblsls")
    For Each rw In tbl.ListRows
        If rw.Range.Cells(1, 1).Value = SaleID Then
            Set sale = New clsSale
            sale.SaleID = rw.Range.Cells(1, 1).Value
            sale.RecordID = rw.Range.Cells(1, 2).Value
            sale.AlbumTitle = rw.Range.Cells(1, 3).Value
            sale.QuantitySold = rw.Range.Cells(1, 4).Value
            sale.SalePrice = rw.Range.Cells(1, 5).Value
            sale.SaleDate = rw.Range.Cells(1, 6).Value
            sale.CustomerID = rw.Range.Cells(1, 7).Value
            Set GetSaleByID = sale
            Exit Function
        End If
    Next rw
End Function

Public Function GetTotalRevenue() As Currency
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim revenue As Currency
    Set ws = ThisWorkbook.Worksheets("Sales")
    Set tbl = ws.ListObjects("tblsls")
    revenue = 0
    For Each rw In tbl.ListRows
        revenue = revenue + _
            (CCur(rw.Range.Cells(1, 5).Value) * _
             CLng(rw.Range.Cells(1, 4).Value))
    Next rw
    GetTotalRevenue = revenue
End Function

Public Sub RecordSale(ByVal sale As clsSale)
    Dim record As clsMusicRecord
    Set record = GetRecordByID(sale.RecordID)
    If record Is Nothing Then
        MsgBox "Inventory record not found."
        Exit Sub
    End If

    If record.Quantity < sale.QuantitySold Then
        MsgBox "Insufficient stock available."
        Exit Sub
    End If

    AddSale sale
    record.Quantity = _
        record.Quantity - sale.QuantitySold
    UpdateRecord record
    MsgBox "Sale Recorded Successfully", vbInformation
End Sub

Public Function GetTotalSalesCount() As Long
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Worksheets("Sales")
    Set tbl = ws.ListObjects("tblsls")
    GetTotalSalesCount = tbl.ListRows.Count
End Function


' test subs
Public Sub TestAddSale()
    Dim sale As clsSale
    Set sale = New clsSale
    sale.SaleID = 1
    sale.RecordID = 1
    sale.AlbumTitle = "Reload"
    sale.QuantitySold = 2
    sale.SalePrice = 20
    sale.SaleDate = Date
    sale.CustomerID = 1001
    AddSale sale
    MsgBox "Sale Added Successfully"
End Sub

Public Sub TestGetSale()
    Dim sale As clsSale
    Set sale = GetSaleByID(1)
    If Not sale Is Nothing Then
        Debug.Print sale.SaleID
        Debug.Print sale.AlbumTitle
        Debug.Print sale.QuantitySold
    Else
        MsgBox "Sale Not Found"
    End If
End Sub

Public Sub TestRevenue()
    Debug.Print _
        FormatCurrency(GetTotalRevenue())
End Sub

Public Sub TestRecordSale()
    Dim sale As clsSale
    Set sale = New clsSale
    sale.SaleID = 3
    sale.RecordID = 1
    sale.AlbumTitle = "Master of Puppets"
    sale.QuantitySold = 2
    sale.SalePrice = 20
    sale.SaleDate = Date
    sale.CustomerID = 1001
    RecordSale sale
End Sub

Public Function GetNextSaleID() As Long
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Worksheets("Sales")
    Set tbl = ws.ListObjects("tblsls")
    If tbl.ListRows.Count = 0 Then
        GetNextSaleID = 1
    Else
        GetNextSaleID = _
            Application.WorksheetFunction.Max( _
                tbl.ListColumns(1).DataBodyRange) + 1
    End If
End Function

