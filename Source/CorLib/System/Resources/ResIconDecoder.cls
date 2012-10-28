VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResIconDecoder"
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
'    Module: ResIconDecoder
'

''
' Decodes an icon from a .RES formatted byte array.
'
' @remarks This takes a byte array formatted like that in a .RES file
' and extracts the Icon picture.
' @see ResourceReader
' @see WinResourceReader
'
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
        Throw Cor.NewArgumentNullException("Key cannot be Nothing.", "Key")
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    If Key.ResourceType <> ResourceTypes.IconResource Then _
        Throw Cor.NewArgumentException("Invalid resource type to decode.", "Key")
        
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
    If mKey Is Nothing Then Exit Function
    
    Set ReturnKey = mKey
    Set ReturnValue = mIcon
    
    Set mKey = Nothing
    Set mIcon = Nothing
    
    GetResource = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateIcon(ByRef Bytes() As Byte) As StdPicture
    Dim hr As Long
    hr = CreateIconFromResource(Bytes(0), cArray.GetLength(Bytes), True, &H30000)
    If hr = 0 Then Exit Function
    
    Dim NewIcon As PictDesc
    With NewIcon
        .cbSizeofStruct = LenB(NewIcon)
        .picType = vbPicTypeIcon
        .hImage = hr
        .xExt = AsLong(Bytes(4))
        .yExt = AsLong(Bytes(8))
    End With
    
    Dim UnkIID As VBGUID
    With UnkIID
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    Dim Unk As IUnknown
    If OleCreatePictureIndirect(NewIcon, UnkIID, True, Unk) = S_OK Then
        Set CreateIcon = Unk
    End If
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
