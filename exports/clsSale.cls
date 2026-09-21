VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSaleID As Long
Private mRecordID As Long
Private mAlbumTitle As String
Private mQuantitySold As Long
Private mSalePrice As Currency
Private mSaleDate As Date
Private mCustomerID As Long

Public Property Get SaleID() As Long
    SaleID = mSaleID
End Property

Public Property Let SaleID(ByVal Value As Long)
    mSaleID = Value
End Property

Public Property Get RecordID() As Long
    RecordID = mRecordID
End Property

Public Property Let RecordID(ByVal Value As Long)
    mRecordID = Value
End Property

Public Property Get AlbumTitle() As String
    AlbumTitle = mAlbumTitle
End Property

Public Property Let AlbumTitle(ByVal Value As String)
    mAlbumTitle = Value
End Property

Public Property Get QuantitySold() As Long
    QuantitySold = mQuantitySold
End Property

Public Property Let QuantitySold(ByVal Value As Long)
    mQuantitySold = Value
End Property

Public Property Get SalePrice() As Currency
    SalePrice = mSalePrice
End Property

Public Property Let SalePrice(ByVal Value As Currency)
    mSalePrice = Value
End Property

Public Property Get SaleDate() As Date
    SaleDate = mSaleDate
End Property

Public Property Let SaleDate(ByVal Value As Date)
    mSaleDate = Value
End Property

Public Property Get CustomerID() As Long
    CustomerID = mCustomerID
End Property

Public Property Let CustomerID(ByVal Value As Long)
    mCustomerID = Value
End Property
