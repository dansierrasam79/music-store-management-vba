VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMusicRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRecordID As Long
Private mAlbumTitle As String
Private mArtist As String
Private mGenre As String
Private mReleaseYear As Long
Private mPurchasePrice As Currency
Private mSellingPrice As Currency
Private mQuantity As Long
Private mDateAdded As Date

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
Public Property Get Artist() As String
    Artist = mArtist
End Property

Public Property Let Artist(ByVal Value As String)
    mArtist = Value
End Property

Public Property Get Genre() As String
    Genre = mGenre
End Property

Public Property Let Genre(ByVal Value As String)
    mGenre = Value
End Property

Public Property Get ReleaseYear() As Long
    ReleaseYear = mReleaseYear
End Property

Public Property Let ReleaseYear(ByVal Value As Long)
    mReleaseYear = Value
End Property

Public Property Get PurchasePrice() As Currency
    PurchasePrice = mPurchasePrice
End Property

Public Property Let PurchasePrice(ByVal Value As Currency)
    mPurchasePrice = Value
End Property

Public Property Get SellingPrice() As Currency
    SellingPrice = mSellingPrice
End Property

Public Property Let SellingPrice(ByVal Value As Currency)
    mSellingPrice = Value
End Property

Public Property Get Quantity() As Long
    Quantity = mQuantity
End Property

Public Property Let Quantity(ByVal Value As Long)
    mQuantity = Value
End Property

Public Property Get DateAdded() As Date
    DateAdded = mDateAdded
End Property

Public Property Let DateAdded(ByVal Value As Date)
    mDateAdded = Value
End Property

Public Function InventoryValue() As Currency
    InventoryValue = mSellingPrice * mQuantity
End Function

Public Function IsLowStock(ByVal Threshold As Long) As Boolean
    IsLowStock = (mQuantity <= Threshold)
End Function
