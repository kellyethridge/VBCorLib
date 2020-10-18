VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CursorResourceDecoder"
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
' Module: CursorResourceDecoder
'

''
' Decodes cursor byte data from a .RES formatted byte array.
'
' @remarks This works exactly like the IconResourceDecoder.
' @see ResourceReader
' @see WinResourceReader
'
Option Explicit
Implements IResourceDecoder

Private Type PictDesc
    cbSizeofStruct  As Long
    picType         As Long
    hImage          As Long
    xExt            As Long
    yExt            As Long
End Type

Private mKey    As ResourceKey
Private mCursor As StdPicture


''
' Decodes an array of bytes into a Cursor StdPicture object.
'
' @param Key The identifier associated with the array of bytes.
' @param Bytes The byte data containing the encoded cursor.
' @return Returns the number of resourced decoded from the byte array. This
' decoder will return no more than 1.
' @see ResourceKey
' @see ResourceTypese
'
Public Function Decode(ByVal Key As ResourceKey, ByRef Bytes() As Byte) As Long
    If Key Is Nothing Then _
        Error.ArgumentNull "Key"
    If CorArray.IsNull(Bytes) Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    If Key.ResourceType <> ResourceTypes.CursorResource Then _
        Error.Argument Argument_InvalidResourceKeyType, "Key"
        
    Set mCursor = CreateCursor(Bytes)
    
    If Not mCursor Is Nothing Then
        Set mKey = Key
        Decode = 1
    End If
End Function

''
' Returns the type of resource this decoder supports.
'
' @return This decoder returns <b>ResourceTypes.CursorResource</b>.
' @see ResourceTypes
'
Public Function GetDecodeType() As Variant
    GetDecodeType = ResourceTypes.CursorResource
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
    Set ReturnValue = mCursor
    
    Set mKey = Nothing
    Set mCursor = Nothing
    
    GetResource = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateCursor(ByRef Bytes() As Byte) As StdPicture
    Dim Handle      As Long
    Dim NewCursor   As PictDesc
    Dim UnkIID      As VBGUID
    Dim Unk         As IUnknown
    
    Handle = CreateIconFromResource(Bytes(LBound(Bytes)), CorArray.Length(Bytes), BOOL_FALSE, &H30000)
    
    If Handle = 0 Then
        Exit Function
    End If
    
    With NewCursor
        .cbSizeofStruct = LenB(NewCursor)
        .picType = vbPicTypeIcon
        .hImage = Handle
        .xExt = AsLong(Bytes(4))
        .yExt = AsLong(Bytes(8))
    End With
    
    With UnkIID
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    If OleCreatePictureIndirect(NewCursor, UnkIID, BOOL_TRUE, Unk) = S_OK Then
        Set CreateCursor = Unk
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

