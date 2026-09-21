Attribute VB_Name = "modReports"
Public Sub GenerateInventoryReport()
    Dim wsReport As Worksheet
    Dim wsInv As Worksheet
    Set wsReport = Worksheets("Reports")
    Set wsInv = Worksheets("Inventory")
    wsReport.Cells.Clear
    wsInv.UsedRange.Copy _
        Destination:=wsReport.Range("A1")
    MsgBox "Inventory Report Generated", vbInformation
End Sub

Public Sub GenerateLowStockReport()
    Dim ws As Worksheet
    Dim results As Collection
    Dim record As clsMusicRecord
    Dim rowNum As Long
    Set ws = Worksheets("Reports")
    ws.Cells.Clear
    ws.Range("A1").Value = "Low Stock Report"
    ws.Range("A3").Value = "Record ID"
    ws.Range("B3").Value = "Album Title"
    ws.Range("C3").Value = "Artist"
    ws.Range("D3").Value = "Quantity"
    Set results = GetLowStockRecords(5)
    rowNum = 4
    For Each record In results
        ws.Cells(rowNum, 1).Value = record.RecordID
        ws.Cells(rowNum, 2).Value = record.AlbumTitle
        ws.Cells(rowNum, 3).Value = record.Artist
        ws.Cells(rowNum, 4).Value = record.Quantity
        rowNum = rowNum + 1
    Next record
    MsgBox "Low Stock Report Generated", vbInformation
End Sub

Public Sub GenerateSalesReport()
    Dim wsReport As Worksheet
    Dim wsSales As Worksheet
    Set wsReport = Worksheets("Reports")
    Set wsSales = Worksheets("Sales")
    wsReport.Cells.Clear
    wsSales.UsedRange.Copy _
        Destination:=wsReport.Range("A1")
    MsgBox "Sales Report Generated", vbInformation
End Sub

Public Sub GenerateRevenueSummary()
    Dim ws As Worksheet
    Dim totalRevenue As Currency
    Dim totalSales As Long
    Dim avgRevenue As Currency
    Set ws = ThisWorkbook.Worksheets("Reports")
    ws.Cells.Clear
    totalRevenue = GetTotalRevenue()
    totalSales = GetTotalSalesCount()

    If totalSales > 0 Then
        avgRevenue = totalRevenue / totalSales
    Else
        avgRevenue = 0
    End If

    ' Report Title
    ws.Range("A1").Value = "Revenue Summary Report"
    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 16
    End With

    ' Report Timestamp
    ws.Range("A3").Value = "Generated On"
    ws.Range("B3").Value = Format(Now, "dd-mmm-yyyy hh:mm:ss")

    ' Revenue Statistics
    ws.Range("A5").Value = "Total Revenue"
    ws.Range("B5").Value = totalRevenue
    ws.Range("A6").Value = "Total Sales"
    ws.Range("B6").Value = totalSales
    ws.Range("A7").Value = "Average Revenue Per Sale"
    ws.Range("B7").Value = avgRevenue

    ' Formatting
    ws.Range("A3:A7").Font.Bold = True
    ws.Range("B5").NumberFormat = "$#,##0.00"
    ws.Range("B7").NumberFormat = "$#,##0.00"
    ws.Columns("A:B").AutoFit
    MsgBox "Revenue Summary Report Generated", vbInformation
End Sub
