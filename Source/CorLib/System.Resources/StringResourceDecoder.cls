VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringResourceDecoder"
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
' Module: StringResourceDecoder
'

''
' Decodes a resource String table containing up to 16 strings.
'
' @see IResourceDecoder
'
Option Explicit
Implements IResourceDecoder

Private mStrings    As Queue
Private mKeys       As Queue


''
' Decodes an array of bytes into a string table containing up to 16 strings.
'
' @param Key The resource identifier for the specific string table.
' @param Bytes The byte data that is decoded to a string table.
' @return The number of strings decoded from the byte array.
' @remarks Up to 16 strings can be decoded from a resource string table.
' @see ResourceKey
'
Public Function Decode(ByVal Key As ResourceKey, ByRef Bytes() As Byte) As Long
    If Key Is Nothing Then _
        Error.ArgumentNull "Key"
    If CorArray.IsNull(Bytes) Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    If Key.ResourceType <> ResourceTypes.StringResource Then _
        Error.Argument Argument_InvalidResourceKeyType, "Key"
        
    Set mStrings = New Queue
    Set mKeys = New Queue
    
    Dim StringId As Long
    ' ResourceName is the string block number. We push out the block number
    ' 4 bits to make room to add the ids of the individual strings within the
    ' block. There are 16 strings in a block, including null strings. The
    ' resulting key for each string in the block will be a combination of the
    ' block number (upper 12 bits) and the string index within the block (lower 4 bits).
    StringId = (CLng(Key.ResourceName) - 1) * vbShift4Bits
    
    Dim ByteIndex       As Long
    Dim MaxByteIndex    As Long
    
    ByteIndex = LBound(Bytes)
    MaxByteIndex = UBound(Bytes)
    
    ' Loop through all the bytes in pairs, as characters,
    ' building the block of 16 strings from the table.
    Do While ByteIndex <= MaxByteIndex
        Dim StringLength As Long
        StringLength = BitConverter.ToInteger(Bytes, ByteIndex)
        ByteIndex = ByteIndex + 2
        
        ' We have a string, so append all the characters.
        If StringLength > 0 Then
            Dim s As String
            s = Encoding.Unicode.GetString(Bytes, ByteIndex, StringLength * vbSizeOfChar)
            ByteIndex = ByteIndex + LenB(s)
            
            mStrings.Enqueue s
            mKeys.Enqueue Cor.NewResourceKey(StringId, ResourceTypes.StringResource, Key.LanguageID)
        End If
        
        StringId = StringId + 1
    Loop
    
    Decode = mStrings.Count
End Function

''
' Returns the type of resource that can be decoded with this decoder.
'
' @return The type of resource that can be decoded.
' @see ResourceTypes
'
Public Function GetDecodeType() As Variant
    GetDecodeType = ResourceTypes.StringResource
End Function


''
' Returns the next available resource in the decoder.
'
' @param ReturnKey This is set to the key that identifies the resource being returned.
' @param ReturnValue This is set to the resource value being returned.
' @return If a resource has been returned, this returns True, otherwise False is returned.
' @remarks Once a resource has been returned, that resource is never returned again.
' @see ResourceKey
'
Public Function GetResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue As Variant) As Boolean
    If mStrings.Count > 0 Then
        Set ReturnKey = mKeys.Dequeue
        ReturnValue = mStrings.Dequeue
        GetResource = True
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
