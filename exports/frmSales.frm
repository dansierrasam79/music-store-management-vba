VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSales 
   Caption         =   "Sales"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "frmSales.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    txtSaleDate.Value = Date
    txtSaleID.Value = GetNextSaleID()
    txtSaleDate.Value = Date
End Sub

Private Sub ClearForm()
    txtSaleID.Value = ""
    txtRecordID.Value = ""
    txtAlbumTitle.Value = ""
    txtQuantitySold.Value = ""
    txtSalePrice.Value = ""
    txtCustomerID.Value = ""
    txtSaleDate.Value = Date
End Sub

Private Sub cmdClear_Click()
    ClearForm
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Function ValidateInputs() As Boolean
    ValidateInputs = False
    If Not IsNumeric(txtSaleID.Value) Then
        MsgBox "Sale ID must be numeric."
        Exit Function
    End If

    If Not IsNumeric(txtRecordID.Value) Then
        MsgBox "Record ID must be numeric."
        Exit Function
    End If

    If Not IsNumeric(txtQuantitySold.Value) Then
        MsgBox "Quantity Sold must be numeric."
        Exit Function
    End If

    If Not IsNumeric(txtSalePrice.Value) Then
        MsgBox "Sale Price must be numeric."
        Exit Function
    End If

    If Not IsDate(txtSaleDate.Value) Then
        MsgBox "Invalid Sale Date."
        Exit Function
    End If

    ValidateInputs = True
End Function

Private Sub cmdRecordSale_Click()
    Dim sale As clsSale
    
    If Not ValidateInputs() Then Exit Sub
    Set sale = New clsSale
    sale.SaleID = CLng(txtSaleID.Value)
    sale.RecordID = CLng(txtRecordID.Value)
    sale.AlbumTitle = txtAlbumTitle.Value
    sale.QuantitySold = CLng(txtQuantitySold.Value)
    sale.SalePrice = CCur(txtSalePrice.Value)
    sale.SaleDate = CDate(txtSaleDate.Value)
    sale.CustomerID = CLng(txtCustomerID.Value)
    RecordSale sale
    ClearForm
End Sub

Private Sub cmdSearchSale_Click()

    Dim sale As clsSale

    If Not IsNumeric(txtSaleID.Value) Then

        MsgBox "Enter a valid Sale ID."
        Exit Sub

    End If

    Set sale = GetSaleByID(CLng(txtSaleID.Value))

    If sale Is Nothing Then

        MsgBox "Sale not found."
        Exit Sub

    End If

    txtRecordID.Value = sale.RecordID
    txtAlbumTitle.Value = sale.AlbumTitle
    txtQuantitySold.Value = sale.QuantitySold
    txtSalePrice.Value = sale.SalePrice
    txtSaleDate.Value = sale.SaleDate
    txtCustomerID.Value = sale.CustomerID

End Sub
Private Sub txtRecordID_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim record As clsMusicRecord
    If Not IsNumeric(txtRecordID.Value) Then Exit Sub
    Set record = GetRecordByID(CLng(txtRecordID.Value))
    If Not record Is Nothing Then
        txtAlbumTitle.Value = record.AlbumTitle
        txtSalePrice.Value = record.SellingPrice
    End If
End Sub
