VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CursorResourceGroupDecoder"
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
' Module: CursorResourceGroupDecoder
'

''
' Decodes a CursorGroup resource.
'
' @remarks This returns a <a href="PictureResourceGroup.html">PictureResourceGroup</a> object when decoded.
'
' @see ResourceReader
' @see WinResourceReader
'
'@Folder("CorLib.System.Resources")
Option Explicit
Implements IResourceDecoder

Private mKey    As ResourceKey
Private mGroup  As PictureResourceGroup


''
' Decodes a byte array into a <a href="PictureResourceGroup.html">PictureResourceGroup</a> that represents a cursor group.
'
' @param Key The identifier for the cursor group.
' @param Bytes The bytes to be decoded into a cursor group.
' @return The number of groups decoded. This will return 0 or 1.
'
Public Function Decode(ByVal Key As ResourceKey, ByRef Bytes() As Byte) As Long
    If Key Is Nothing Then _
        Error.ArgumentNull "Key"
    If CorArray.IsNull(Bytes) Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    If Key.ResourceType <> ResourceTypes.GroupCursor Then _
        Error.Argument Argument_InvalidResourceKeyType, "Key"

    Set mKey = Key
    Set mGroup = New PictureResourceGroup
    mGroup.InitFromBytes Bytes
    Decode = 1
End Function

''
' Returns the type of resource this decoder supports.
'
' @return This decoder returns <a href="ResourceKey.ResourceTypes.html">ResourceTypes.GroupCursor</a>.
' @see ResourceTypes
'
Public Function GetDecodeType() As Variant
    GetDecodeType = ResourceTypes.GroupCursor
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
    Set ReturnValue = mGroup
    Set mKey = Nothing
    Set mGroup = Nothing
    GetResource = True
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

