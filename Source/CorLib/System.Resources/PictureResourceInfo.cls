VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PictureResourceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2018 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: PictureResourceInfo
'

''
' Represents a picture resource that is part of a group of pictures.
'
' @see ResPictureGroup
'
Option Explicit
Implements IObject

Private mWidth          As Long
Private mHeight         As Long
Private mColorCount     As Long
Private mPlanes         As Long
Private mBitCount       As Long
Private mBytesInRes     As Long
Private mResourceID     As Long
Private mResourceType   As ResourceTypes


''
' Returns the width of the picture.
'
' @return The picture width in pixels.
'
Public Property Get Width() As Long
    Width = mWidth
End Property

''
' Returns the height of the picture.
'
' @return The picture height in pixels.
'
Public Property Get Height() As Long
    Height = mHeight
End Property

''
' Returns the number of colors used in the picture.
'
' @return The number of colors used.
'
Public Property Get Colors() As Long
    Colors = mColorCount
End Property

''
' Returns the number of bits used to represent the color range.
'
' @return The number of bits used to represent the color range.
'
Public Property Get BitCount() As Long
    BitCount = mBitCount
End Property

''
' The numeric identifier of the associated picture.
'
' @return The numeric identifier.
'
Public Property Get ResourceId() As Long
    ResourceId = mResourceID
End Property

''
' Returns the number of planes. This returns 1.
'
' @return The number of planes.
'
Public Property Get Planes() As Long
    Planes = mPlanes
End Property

''
' Returns the number of bytes in this resource.
'
' @return The number of bytes in this resource.
'
Public Property Get Size() As Long
    Size = mBytesInRes
End Property

''
' Returns the type of group this resource belongs to.
'
' @return The type of group this resource belong to.
'
Public Property Get ResourceType() As ResourceTypes
    ResourceType = mResourceType
End Property

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to compare to this instance.
' @return Returns True if the value is equal to this instance, False otherwise.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' @return A hashcode value.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' @return A string representation of this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function Parse(ByVal ResType As PictureGroupTypes, ByVal Index As Long, ByRef Bytes() As Byte) As Long
    If ResType = IconGroup Then
        mWidth = Bytes(Index)
        mHeight = Bytes(Index + 1)
        mColorCount = Bytes(Index + 2)
        mResourceType = IconResource
    Else
        mWidth = AsWord(Bytes(Index))
        mHeight = AsWord(Bytes(Index + 2))
        mResourceType = CursorResource
    End If
    
    Index = Index + 4
    
    mPlanes = BitConverter.ToInteger(Bytes, Index)
    Index = Index + 2
    
    mBitCount = BitConverter.ToInteger(Bytes, Index)
    Index = Index + 2
    
    If ResType = CursorGroup Then
        mColorCount = Powers(mBitCount)
    End If
    
    mBytesInRes = BitConverter.ToLong(Bytes, Index)
    Index = Index + 4
    
    mResourceID = BitConverter.ToInteger(Bytes, Index)
    Parse = Index + 2
End Function

Friend Sub InitFromReader(ByVal ResourceType As PictureGroupTypes, ByVal Reader As BinaryReader)
    mResourceType = ResourceType
    mWidth = Reader.ReadByte
    mHeight = Reader.ReadByte
    mColorCount = Reader.ReadInt16
    mPlanes = Reader.ReadInt16
    mBitCount = Reader.ReadInt16
    mBytesInRes = Reader.ReadInt32
    mResourceID = Reader.ReadInt16
End Sub

Friend Sub Init(ByVal ResourceId As Long, ByVal Resource As IPicture, Optional ByVal ResourceType As PictureGroupTypes)
    Dim Info    As ICONINFO
    Dim Encoder As IResourceEncoder
    Dim Data()  As Byte
    Dim Header  As BITMAPINFOHEADER
    
    If Resource Is Nothing Then _
        Error.ArgumentNull "Resource"
    
    mResourceID = ResourceId
        
    If GetIconInfo(Resource.Handle, Info) = BOOL_FALSE Then _
        ResourceTypeError
    
    If Info.fIcon Then
        Set Encoder = New IconResourceEncoder
        mResourceType = IconResource
    Else
        Set Encoder = New CursorResourceEncoder
        mResourceType = CursorResource
    End If
    
    If Not Encoder.Encode(Resource, ResourceId) Then _
        ResourceTypeError
    If Not Encoder.GetEncodedResource(Nothing, Data) Then _
        ResourceTypeError
       
    CopyMemory Header, Data(0), Len(Header)
    
    With Header
        mWidth = .biWidth
        mHeight = .biHeight
        mColorCount = Powers(.biBitCount)
        mPlanes = .biPlanes
        mBitCount = .biBitCount
        mBytesInRes = CorArray.Length(Data)
    End With
End Sub

Private Sub ResourceTypeError()
    Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_MustBeOfType, "Icon or Cursor"), "Resource")
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
