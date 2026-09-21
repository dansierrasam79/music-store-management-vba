VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReports 
   Caption         =   "UserForm1"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   OleObjectBlob   =   "frmReports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInventoryReport_Click()
    GenerateInventoryReport
End Sub

Private Sub cmdLowStockReport_Click()
    GenerateLowStockReport
End Sub

Private Sub cmdSalesReport_Click()
    GenerateSalesReport
End Sub

Private Sub cmdRevenueReport_Click()
    GenerateRevenueSummary
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub
