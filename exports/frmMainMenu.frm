VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMainMenu 
   Caption         =   "Music Store Management Store"
   ClientHeight    =   2340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmMainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDashboard_Click()
    frmDashboard.Show
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdInventory_Click()
    frmInventory.Show
End Sub

Private Sub cmdReports_Click()
    frmReports.Show
End Sub

Private Sub cmdSales_Click()
    frmSales.Show
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

