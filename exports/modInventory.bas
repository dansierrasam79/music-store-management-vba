Attribute VB_Name = "modInventory"
Option Explicit

Public Sub AddRecord(ByVal record As clsMusicRecord)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Set ws = ThisWorkbook.Worksheets("Inventory")
    Set tbl = ws.ListObjects("tblInv")
    Set newRow = tbl.ListRows.Add

    With newRow.Range
        .Cells(1, 1).Value = record.RecordID
        .Cells(1, 2).Value = record.AlbumTitle
        .Cells(1, 3).Value = record.Artist
        .Cells(1, 4).Value = record.Genre
        .Cells(1, 5).Value = record.ReleaseYear
        .Cells(1, 6).Value = record.PurchasePrice
        .Cells(1, 7).Value = record.SellingPrice
        .Cells(1, 8).Value = record.Quantity
        .Cells(1, 9).Value = record.DateAdded
    End With
End Sub

Public Function GetRecordByID(ByVal RecordID As Long) As clsMusicRecord
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim record As clsMusicRecord

    Set ws = ThisWorkbook.Worksheets("Inventory")
    Set tbl = ws.ListObjects("tblInv")
    For Each rw In tbl.ListRows
        If rw.Range.Cells(1, 1).Value = RecordID Then
            Set record = New clsMusicRecord
            record.RecordID = rw.Range.Cells(1, 1).Value
            record.AlbumTitle = rw.Range.Cells(1, 2).Value
            record.Artist = rw.Range.Cells(1, 3).Value
            record.Genre = rw.Range.Cells(1, 4).Value
            record.ReleaseYear = rw.Range.Cells(1, 5).Value
            record.PurchasePrice = rw.Range.Cells(1, 6).Value
            record.SellingPrice = rw.Range.Cells(1, 7).Value
            record.Quantity = rw.Range.Cells(1, 8).Value
            record.DateAdded = rw.Range.Cells(1, 9).Value
            Set GetRecordByID = record
            Exit Function
        End If
    Next rw
End Function

Public Sub UpdateRecord(ByVal record As clsMusicRecord)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Set ws = ThisWorkbook.Worksheets("Inventory")
    Set tbl = ws.ListObjects("tblInv")
    For Each rw In tbl.ListRows
        If rw.Range.Cells(1, 1).Value = record.RecordID Then
            rw.Range.Cells(1, 2).Value = record.AlbumTitle
            rw.Range.Cells(1, 3).Value = record.Artist
            rw.Range.Cells(1, 4).Value = record.Genre
            rw.Range.Cells(1, 5).Value = record.ReleaseYear
            rw.Range.Cells(1, 6).Value = record.PurchasePrice
            rw.Range.Cells(1, 7).Value = record.SellingPrice
            rw.Range.Cells(1, 8).Value = record.Quantity
            rw.Range.Cells(1, 9).Value = record.DateAdded
            Exit For
        End If
    Next rw
End Sub

Public Sub DeleteRecord(ByVal RecordID As Long)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim Deleted As Boolean
    Set ws = ThisWorkbook.Worksheets("Inventory")
    Set tbl = ws.ListObjects("tblInv")
    Deleted = False
    For i = tbl.ListRows.Count To 1 Step -1
        If tbl.ListRows(i).Range.Cells(1, 1).Value = RecordID Then
            tbl.ListRows(i).Delete
            Deleted = True
            Exit For
        End If
    Next i
    
    If Deleted Then
        MsgBox "Record Deleted Successfully"
    Else
        MsgBox "Record Not Found"
    End If
End Sub

Public Function SearchRecordsByArtist(ByVal ArtistName As String) As Collection
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim record As clsMusicRecord
    Dim results As Collection

    Set ws = ThisWorkbook.Worksheets("Inventory")
    Set tbl = ws.ListObjects("tblInv")
    Set results = New Collection
    For Each rw In tbl.ListRows
        If LCase(Trim(rw.Range.Cells(1, 3).Value)) = _
           LCase(Trim(ArtistName)) Then
            Set record = New clsMusicRecord
            record.RecordID = rw.Range.Cells(1, 1).Value
            record.AlbumTitle = rw.Range.Cells(1, 2).Value
            record.Artist = rw.Range.Cells(1, 3).Value
            record.Genre = rw.Range.Cells(1, 4).Value
            record.ReleaseYear = rw.Range.Cells(1, 5).Value
            record.PurchasePrice = rw.Range.Cells(1, 6).Value
            record.SellingPrice = rw.Range.Cells(1, 7).Value
            record.Quantity = rw.Range.Cells(1, 8).Value
            record.DateAdded = rw.Range.Cells(1, 9).Value
            results.Add record
        End If
    Next rw
    Set SearchRecordsByArtist = results
End Function

Public Function SearchRecordsByAlbum(ByVal AlbumName As String) As Collection
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim record As clsMusicRecord
    Dim results As Collection

    Set ws = ThisWorkbook.Worksheets("Inventory")
    Set tbl = ws.ListObjects("tblInv")
    Set results = New Collection
    For Each rw In tbl.ListRows
        If InStr(1, _
                 LCase(rw.Range.Cells(1, 2).Value), _
                 LCase(AlbumName)) > 0 Then

            Set record = New clsMusicRecord
            record.RecordID = rw.Range.Cells(1, 1).Value
            record.AlbumTitle = rw.Range.Cells(1, 2).Value
            record.Artist = rw.Range.Cells(1, 3).Value
            record.Genre = rw.Range.Cells(1, 4).Value
            record.ReleaseYear = rw.Range.Cells(1, 5).Value
            record.PurchasePrice = rw.Range.Cells(1, 6).Value
            record.SellingPrice = rw.Range.Cells(1, 7).Value
            record.Quantity = rw.Range.Cells(1, 8).Value
            record.DateAdded = rw.Range.Cells(1, 9).Value
            results.Add record
        End If
    Next rw
    Set SearchRecordsByAlbum = results
End Function

Public Function SearchRecordsByGenre(ByVal GenreName As String) As Collection
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim record As clsMusicRecord
    Dim results As Collection

    Set ws = ThisWorkbook.Worksheets("Inventory")
    Set tbl = ws.ListObjects("tblInv")
    Set results = New Collection
    For Each rw In tbl.ListRows
        If InStr(1, _
                 LCase(rw.Range.Cells(1, 4).Value), _
                 LCase(GenreName)) > 0 Then

            Set record = New clsMusicRecord
            record.RecordID = rw.Range.Cells(1, 1).Value
            record.AlbumTitle = rw.Range.Cells(1, 2).Value
            record.Artist = rw.Range.Cells(1, 3).Value
            record.Genre = rw.Range.Cells(1, 4).Value
            record.ReleaseYear = rw.Range.Cells(1, 5).Value
            record.PurchasePrice = rw.Range.Cells(1, 6).Value
            record.SellingPrice = rw.Range.Cells(1, 7).Value
            record.Quantity = rw.Range.Cells(1, 8).Value
            record.DateAdded = rw.Range.Cells(1, 9).Value
            results.Add record
        End If
    Next rw
    Set SearchRecordsByGenre = results
End Function

Public Function GetLowStockRecords(ByVal Threshold As Long) As Collection
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim record As clsMusicRecord
    Dim results As Collection

    Set ws = ThisWorkbook.Worksheets("Inventory")
    Set tbl = ws.ListObjects("tblInv")
    Set results = New Collection
    For Each rw In tbl.ListRows
        If CLng(rw.Range.Cells(1, 8).Value) <= Threshold Then
            Set record = New clsMusicRecord
            record.RecordID = rw.Range.Cells(1, 1).Value
            record.AlbumTitle = rw.Range.Cells(1, 2).Value
            record.Artist = rw.Range.Cells(1, 3).Value
            record.Genre = rw.Range.Cells(1, 4).Value
            record.ReleaseYear = rw.Range.Cells(1, 5).Value
            record.PurchasePrice = rw.Range.Cells(1, 6).Value
            record.SellingPrice = rw.Range.Cells(1, 7).Value
            record.Quantity = rw.Range.Cells(1, 8).Value
            record.DateAdded = rw.Range.Cells(1, 9).Value
            results.Add record
        End If
    Next rw
    Set GetLowStockRecords = results
End Function

Public Function GetInventoryValue() As Currency
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim TotalValue As Currency
    Set ws = ThisWorkbook.Worksheets("Inventory")
    Set tbl = ws.ListObjects("tblInv")
    TotalValue = 0

    For Each rw In tbl.ListRows
        TotalValue = TotalValue + _
            (CCur(rw.Range.Cells(1, 7).Value) * _
             CLng(rw.Range.Cells(1, 8).Value))
    Next rw
    GetInventoryValue = TotalValue
End Function

Public Function GetRecordCount() As Long
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Worksheets("Inventory")
    Set tbl = ws.ListObjects("tblInv")
    GetRecordCount = tbl.ListRows.Count
End Function


' test functions
Public Sub TestAddRecord()
    Dim record As clsMusicRecord
    Set record = New clsMusicRecord
    record.RecordID = 5
    record.AlbumTitle = "Metallica"
    record.Artist = "Metallica"
    record.Genre = "Metal"
    record.ReleaseYear = 1991
    record.PurchasePrice = 10
    record.SellingPrice = 20
    record.Quantity = 5
    record.DateAdded = Date
    AddRecord record
    MsgBox "Record Added Successfully"
End Sub

Public Sub TestGetRecord()
    Dim record As clsMusicRecord
    Set record = GetRecordByID(1)
    If Not record Is Nothing Then
        Debug.Print record.RecordID
        Debug.Print record.AlbumTitle
        Debug.Print record.Artist
        Debug.Print record.Genre
        Debug.Print record.ReleaseYear
    Else
        MsgBox "Record not found"
    End If
End Sub

Public Sub TestUpdateRecord()
    Dim record As clsMusicRecord
    Set record = GetRecordByID(1)
    If record Is Nothing Then
        MsgBox "Record not found"
        Exit Sub
    End If
    record.Quantity = 10
    UpdateRecord record
    MsgBox "Record Updated"
End Sub

Public Sub TestDeleteRecord()
    DeleteRecord 2
End Sub

Public Sub TestSearchRecordsByArtist()
    Dim results As Collection
    Dim record As clsMusicRecord
    Set results = SearchRecordsByArtist("Metallica")
    
    Debug.Print "Records Found: " & results.Count
    For Each record In results
        Debug.Print record.AlbumTitle
    Next record
End Sub

Public Sub TestSearchByAlbum()
    Dim results As Collection
    Dim record As clsMusicRecord
    Set results = SearchRecordsByAlbum("Master")
    Debug.Print "Albums Found: " & results.Count
    For Each record In results
        Debug.Print record.AlbumTitle
    Next record
End Sub

Public Sub TestSearchByGenre()
    Dim results As Collection
    Dim record As clsMusicRecord
    Set results = SearchRecordsByGenre("Metal")
    Debug.Print "Records Found: " & results.Count
    For Each record In results
        Debug.Print record.AlbumTitle
    Next record
End Sub

Public Sub TestLowStockRecords()
    Dim results As Collection
    Dim record As clsMusicRecord
    Set results = GetLowStockRecords(5)
    Debug.Print "Low Stock Records: " & results.Count
    For Each record In results
        Debug.Print record.AlbumTitle & _
                    " - Qty: " & record.Quantity
    Next record
End Sub

Public Sub TestInventoryValue()
    Debug.Print "Inventory Value: " & _
                FormatCurrency(GetInventoryValue())
End Sub
