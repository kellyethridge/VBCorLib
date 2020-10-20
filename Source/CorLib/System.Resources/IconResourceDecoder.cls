VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IconResourceDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: IconResourceDecoder
'

''
' Decodes an icon from a .RES formatted byte array.
'
' @remarks This takes a byte array formatted like that in a .RES file and extracts the Icon picture.
' @see ResourceReader
' @see WinResourceReader
'
'@Folder("CorLib.System.Resources")
Option Explicit
Implements IResourceDecoder

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private mKey    As ResourceKey
Private mIcon   As StdPicture


''
' Decodes a byte array into a <b>StdPicture</b> that represents an icon.
'
' @param Key The identifier for the icon.
' @param Bytes The bytes to be decoded into an icon.
' @return The number of icons decoded. This will return 0 or 1.
'
Public Function Decode(ByVal Key As ResourceKey, ByRef Bytes() As Byte) As Long
    If Key Is Nothing Then _
        Error.ArgumentNull "Key"
    If CorArray.IsNull(Bytes) Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    If Key.ResourceType <> ResourceTypes.IconResource Then _
        Error.Argument Argument_InvalidResourceKeyType, "Key"
    
    Set mIcon = CreateIcon(Bytes)
    If Not mIcon Is Nothing Then
        Set mKey = Key
        Decode = 1
    End If
End Function

''
' Returns the type of resource this decoder supports.
'
' @return This decoder returns <b>ResourceTypes.IconResource</b>.
' @see ResourceTypes
'
Public Function GetDecodeType() As Variant
    GetDecodeType = ResourceTypes.IconResource
End Function

''
' Returns the next available resource in the decoder.
'
' @param ReturnKey This is set to the key that identifies the resource being returned.
' @param ReturnValue This is set to the resource value being returned.
' @return If a resource has been returned, this returns True, otherwise False is returned.
' @remarks Once a resource has been returned, that resource is never returned again.
'
Public Function GetResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue As Variant) As Boolean
    If mKey Is Nothing Then
        Exit Function
    End If
    
    Set ReturnKey = mKey
    Set ReturnValue = mIcon
    
    Set mKey = Nothing
    Set mIcon = Nothing
    
    GetResource = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateIcon(ByRef Bytes() As Byte) As StdPicture
    Dim Handle  As Long
    Dim NewIcon As PictDesc
    Dim UnkIID  As VBGUID
    Dim Unk     As IUnknown
    
    Handle = CreateIconFromResource(Bytes(LBound(Bytes)), CorArray.Length(Bytes), BOOL_TRUE, &H30000)
    
    If Handle = 0 Then
        Exit Function
    End If
    
    With NewIcon
        .cbSizeofStruct = LenB(NewIcon)
        .picType = vbPicTypeIcon
        .hImage = Handle
        .xExt = AsLong(Bytes(4))
        .yExt = AsLong(Bytes(8))
    End With
    
    With UnkIID
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    If OleCreatePictureIndirect(NewIcon, UnkIID, BOOL_TRUE, Unk) = S_OK Then
        Set CreateIcon = Unk
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceDecoder
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResourceDecoder_Decode(ByVal Key As ResourceKey, Bytes() As Byte) As Long
    IResourceDecoder_Decode = Decode(Key, Bytes)
End Function

Private Function IResourceDecoder_GetDecodeType() As Variant
    IResourceDecoder_GetDecodeType = GetDecodeType
End Function

Private Function IResourceDecoder_GetResource(ReturnKey As ResourceKey, ReturnValue As Variant) As Boolean
    IResourceDecoder_GetResource = GetResource(ReturnKey, ReturnValue)
End Function
