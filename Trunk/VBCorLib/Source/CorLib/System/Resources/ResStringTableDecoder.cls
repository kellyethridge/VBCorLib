VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResStringTableDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: ResStringTableDecoder
'

''
' Decodes a resource String table containing up to 16 strings.
'
' @see IResourceDecoder
'
Option Explicit
Implements IResourceDecoder

Private mStrings(15)    As String
Private mKeys(15)       As ResourceKey
Private mIndex          As Long
Private mKey            As ResourceKey
Private mStringBuilder  As New StringBuilder


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
        Throw Cor.NewArgumentNullException("Key")
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException("Bytes", Environment.GetResourceString(ArgumentNull_Array))
    If Key.ResourceType <> ResourceTypes.StringResource Then _
        Throw Cor.NewArgumentException("Invalid resource type to decode.", "Key")
        
    Set mKey = Key
    
    Dim StringID As Long
    ' ResourceName should be a number.
    StringID = (CLng(mKey.ResourceName) - 1) * 16
    
    ' we need to reset this because this decoder may
    ' decode multiple times.
    mIndex = 0
    
    Dim MaxByteIndex As Long
    MaxByteIndex = UBound(Bytes)
    
    Dim Count As Long
    Dim ByteIndex As Long
    
    ' Loop through all the bytes in pairs, as characters,
    ' building the block of 16 strings from the table.
    Do While ByteIndex <= MaxByteIndex
        Dim StringLength As Long
        StringLength = BitConverter.ToInteger(Bytes, ByteIndex)
        ByteIndex = ByteIndex + 2
        
        ' We have a string, so append all the characters.
        If StringLength > 0 Then
            mStringBuilder.Length = 0
            
            Dim i As Long
            For i = 1 To StringLength
                Call mStringBuilder.AppendChar(BitConverter.ToInteger(Bytes, ByteIndex))
                ByteIndex = ByteIndex + 2
            Next i
            mStrings(mIndex) = mStringBuilder.ToString
            Set mKeys(mIndex) = Cor.NewResourceKey(StringID, ResourceTypes.StringResource, mKey.LanguageID)
            Count = Count + 1
        Else
            mStrings(mIndex) = vbNullString
            Set mKeys(mIndex) = Nothing
        End If
        mIndex = mIndex + 1
        StringID = StringID + 1
    Loop
    
    Decode = Count
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
    Do While mIndex > 0
        mIndex = mIndex - 1
        If Not mKeys(mIndex) Is Nothing Then
            Set ReturnKey = mKeys(mIndex)
            ReturnValue = mStrings(mIndex)
            GetResource = True
            Exit Function
        End If
    Loop
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceDecoder Interface
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
