VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard 
   Caption         =   "Dashboard"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   OleObjectBlob   =   "frmDashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub RefreshDashboard()
    Dim lowStock As Collection
    Dim record As clsMusicRecord
    lblInventoryValue.Caption = _
        FormatCurrency(GetInventoryValue())
    lstLowStock.Clear
    Set lowStock = GetLowStockRecords(5)
    For Each record In lowStock
        lstLowStock.AddItem _
        record.AlbumTitle & _
        " (" & record.Quantity & ")"
    Next record
    lblRecordCount.Caption = GetRecordCount()
    lblRevenue.Caption = _
    FormatCurrency(GetTotalRevenue())
    lblSales.Caption = _
    GetTotalSalesCount()
    lblLastRefresh.Caption = _
        Format(Now, "dd-mmm-yyyy hh:mm")
End Sub

Private Sub UserForm_Initialize()
    RefreshDashboard
End Sub

Private Sub cmdRefresh_Click()
    RefreshDashboard
End Sub

Private Sub cmdBack_Click()
    Me.Hide
    'frmDashboard.Show
    'Me.Show
End Sub

