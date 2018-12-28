VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakePictureResourceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements PictureResourceInfo

Public BitCount     As Long
Public Colors       As Long
Public Height       As Long
Public Planes       As Long
Public ResourceType As CorLib.ResourceTypes
Public ResourceId   As Long
Public Size         As Long
Public Width        As Long

Private Property Get PictureResourceInfo_BitCount() As Long
    PictureResourceInfo_BitCount = BitCount
End Property

Private Property Get PictureResourceInfo_Colors() As Long
    PictureResourceInfo_Colors = Colors
End Property

Private Function PictureResourceInfo_Equals(Value As Variant) As Boolean

End Function

Private Function PictureResourceInfo_GetHashCode() As Long

End Function

Private Property Get PictureResourceInfo_Height() As Long
    PictureResourceInfo_Height = Height
End Property

Private Property Get PictureResourceInfo_Planes() As Long
    PictureResourceInfo_Planes = Planes
End Property

Private Property Get PictureResourceInfo_ResourceId() As Long
    PictureResourceInfo_ResourceId = ResourceId
End Property

Private Property Get PictureResourceInfo_ResourceType() As CorLib.ResourceTypes
    PictureResourceInfo_ResourceType = ResourceType
End Property

Private Property Get PictureResourceInfo_Size() As Long
    PictureResourceInfo_Size = Size
End Property

Private Function PictureResourceInfo_ToString() As String

End Function

Private Property Get PictureResourceInfo_Width() As Long
    PictureResourceInfo_Width = Width
End Property
