VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInventory 
   Caption         =   "Inventory Management"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10095
   OleObjectBlob   =   "frmInventory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClearForm()
    txtRecordID.Value = ""
    txtAlbumTitle.Value = ""
    txtArtist.Value = ""
    txtGenre.Value = ""
    txtReleaseYear.Value = ""
    txtPurchasePrice.Value = ""
    txtSellingPrice.Value = ""
    txtQuantity.Value = ""
    txtDateAdded.Value = Date
End Sub

Private Sub cmdClear_Click()
    ClearForm
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdAdd_Click()
    If Not ValidateInputs() Then Exit Sub
    Dim record As clsMusicRecord
    Dim existingRecord As clsMusicRecord
    Set existingRecord = _
    GetRecordByID(CLng(txtRecordID.Value))
    If Not existingRecord Is Nothing Then
        MsgBox "Record ID already exists."
        Exit Sub
    End If
    Set record = New clsMusicRecord
    record.RecordID = CLng(txtRecordID.Value)
    record.AlbumTitle = txtAlbumTitle.Value
    record.Artist = txtArtist.Value
    record.Genre = txtGenre.Value
    record.ReleaseYear = CLng(txtReleaseYear.Value)
    record.PurchasePrice = CCur(txtPurchasePrice.Value)
    record.SellingPrice = CCur(txtSellingPrice.Value)
    record.Quantity = CLng(txtQuantity.Value)
    record.DateAdded = CDate(txtDateAdded.Value)
    AddRecord record
    MsgBox "Record Added Successfully", vbInformation
    ClearForm
End Sub

Private Sub cmdSearch_Click()
    Dim record As clsMusicRecord
    If Trim(txtRecordID.Value) = "" Then
        MsgBox "Please enter a Record ID."
    Exit Sub
    End If
    Set record = GetRecordByID(CLng(txtRecordID.Value))
    If record Is Nothing Then
        MsgBox "Record Not Found"
        Exit Sub
    End If
    txtAlbumTitle.Value = record.AlbumTitle
    txtArtist.Value = record.Artist
    txtGenre.Value = record.Genre
    txtReleaseYear.Value = record.ReleaseYear
    txtPurchasePrice.Value = record.PurchasePrice
    txtSellingPrice.Value = record.SellingPrice
    txtQuantity.Value = record.Quantity
    txtDateAdded.Value = record.DateAdded
End Sub

Private Sub cmdUpdate_Click()
    Dim record As clsMusicRecord
    Set record = New clsMusicRecord
    If Trim(txtReleaseYear.Value) = "" Then
        MsgBox "Release Year is required."
        txtReleaseYear.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtReleaseYear.Value) Then
        MsgBox "Release Year must be numeric."
        txtReleaseYear.SetFocus
        Exit Sub
    End If
    record.RecordID = CLng(txtRecordID.Value)
    record.AlbumTitle = txtAlbumTitle.Value
    record.Artist = txtArtist.Value
    record.Genre = txtGenre.Value
    record.ReleaseYear = CLng(txtReleaseYear.Value)
    record.PurchasePrice = Val(txtPurchasePrice.Value)
    record.SellingPrice = Val(txtSellingPrice.Value)
    record.Quantity = Val(txtQuantity.Value)
    record.DateAdded = CDate(txtDateAdded.Value)
    UpdateRecord record
    MsgBox "Record Updated Successfully", vbInformation
End Sub

Private Sub cmdDelete_Click()
    If MsgBox( _
        "Delete Record?", _
        vbYesNo + vbQuestion) = vbYes Then
        DeleteRecord CLng(txtRecordID.Value)
        ClearForm
        MsgBox "Record Deleted Successfully", vbInformation
    End If
End Sub

Private Sub UserForm_Initialize()
    txtDateAdded.Value = Date
End Sub

Private Function ValidateInputs() As Boolean
    ValidateInputs = False
    If Trim(txtRecordID.Value) = "" Then
        MsgBox "Record ID is required."
        txtRecordID.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtRecordID.Value) Then
        MsgBox "Record ID must be numeric."
        txtRecordID.SetFocus
        Exit Function
    End If
    If Trim(txtAlbumTitle.Value) = "" Then
        MsgBox "Album Title is required."
        txtAlbumTitle.SetFocus
        Exit Function
    End If

    If Trim(txtArtist.Value) = "" Then
        MsgBox "Artist is required."
        txtArtist.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtReleaseYear.Value) Then
        MsgBox "Release Year must be numeric."
        txtReleaseYear.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtPurchasePrice.Value) Then
        MsgBox "Purchase Price must be numeric."
        txtPurchasePrice.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtSellingPrice.Value) Then
        MsgBox "Selling Price must be numeric."
        txtSellingPrice.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtQuantity.Value) Then
        MsgBox "Quantity must be numeric."
        txtQuantity.SetFocus
        Exit Function
    End If
    
    If Not IsDate(txtDateAdded.Value) Then
        MsgBox "Date Added is invalid."
        txtDateAdded.SetFocus
        Exit Function
    End If

    ValidateInputs = True
End Function

Private Sub txtRecordID_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim record As clsMusicRecord

    If Not IsNumeric(txtRecordID.Value) Then Exit Sub

    Set record = GetRecordByID(CLng(txtRecordID.Value))

    If record Is Nothing Then Exit Sub

    txtAlbumTitle.Value = record.AlbumTitle
    txtArtist.Value = record.Artist
    txtGenre.Value = record.Genre
    txtReleaseYear.Value = record.ReleaseYear
    txtPurchasePrice.Value = record.PurchasePrice
    txtSellingPrice.Value = record.SellingPrice
    txtQuantity.Value = record.Quantity
    txtDateAdded.Value = record.DateAdded

End Sub
