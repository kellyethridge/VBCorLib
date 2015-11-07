VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResIconGroupEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: ResIconGroupEncoder
'

''
' Encodes a <b>ResPictureGroup</b> object to a .RES byte array.
'
' @see ResourceWriter
'
Option Explicit
Implements IResourceEncoder

Private Const SIZEOF_NEWHEADER  As Long = 6
Private Const SIZEOF_RESDIR     As Long = 14
Private Const RES_ICON          As Long = 1
Private Const RES_CURSOR        As Long = 2

Private mResources  As New Hashtable
Private mIterator   As IDictionaryEnumerator



''
' Encodes a <b>ResPictureGroup</b> object.
'
' @param Value The Icon StdPicture to be encodes.
' @param ResourceName The identifier for the resource value. This can be a String or Number.
' @param ResourceType Identifies the type of resource this is. If this is not supplied, then
' the type is derived from the value.
' @param LanguageID The LocaleID the value is associated with.
' @return Returns True if the encoder was able to encode the value, False otherwise.
' @remarks Any value can be passed in. No exception is thrown. If the value is not
' a bitmap then the function returns False.
' <p>The <b>ResourceName</b> can be a String or a Numeric value.</p>
' <p>The LanguageID can be a <b>CultureInfo</b> object, a culture name, or an LCID. If
' the LanguageID is not supplied, then the current culture is used.</p>
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean
    If IsMissing(ResourceType) Then
        If Not IsObject(Value) Then Exit Function
        If Value Is Nothing Then Exit Function
        If Not TypeOf Value Is ResPictureGroup Then Exit Function
    Else
        If Not IsInteger(ResourceType) Then Exit Function
        If ResourceType <> ResourceTypes.GroupIcon Then Exit Function
    End If
    
    Dim Group As ResPictureGroup
    Set Group = Value
    
    If Group.GroupType <> IconGroup Then Exit Function
    If Group.Count = 0 Then Exit Function
    
    Dim Bytes() As Byte
    ReDim Bytes(0 To SIZEOF_NEWHEADER + Group.Count * SIZEOF_RESDIR - 1)
    
    AsWord(Bytes(2)) = RES_ICON '
    AsWord(Bytes(4)) = Group.Count
    
    Dim Index As Long
    Index = SIZEOF_NEWHEADER
    
    Dim i As Long
    For i = 0 To Group.Count - 1
        With Group(i)
            Bytes(Index) = .Width
            Index = Index + 1
            
            Bytes(Index) = .Height
            Index = Index + 1
            
            Bytes(Index) = .Colors
            Index = Index + 2   ' skip byte 4.
            
            AsWord(Bytes(Index)) = .Planes
            Index = Index + 2
            
            AsWord(Bytes(Index)) = .bitCount
            Index = Index + 2
            
            AsLong(Bytes(Index)) = .Size
            Index = Index + 4
            
            AsWord(Bytes(Index)) = AsWord(.ResourceId)
            Index = Index + 2
        End With
    Next i
    
    Dim Key As ResourceKey
    Set Key = Cor.NewResourceKey(ResourceName, ResourceTypes.GroupIcon, GetLanguageID(LanguageID))
    
    Call mResources.Add(Key, Bytes)
    Encode = True
End Function

''
' Retrieves the next encoded resource in the encoder.
'
' @param ReturnKey This is set to a <b>ResourceKey</b> object that
' uniquely identify the encoded resource.
' @param ReturnValue This will be set to a byte array.
' @return Returns True if an encoded resource was returned, False otherwise.
' @remarks This function can be called multiple times to retrieve multiple resources.
' As long as a resource is returned, True will be returned, otherwise False will be returned.
'
Public Function GetEncodedResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue As Variant) As Boolean
    If mIterator Is Nothing Then
        Set mIterator = mResources.GetEnumerator
    End If
    
    If mIterator.MoveNext Then
        Set ReturnKey = mIterator.Key
        ReturnValue = mIterator.Value
        GetEncodedResource = True
    End If
End Function

''
' Releases all currently encoded values.
'
Public Sub Reset()
    Call mResources.Clear
    Set mIterator = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceEncoder Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResourceEncoder_Encode(Value As Variant, ResourceName As Variant, Optional ResourceType As Variant, Optional LanguageID As Variant) As Boolean
    IResourceEncoder_Encode = Encode(Value, ResourceName, ResourceType, LanguageID)
End Function

Private Function IResourceEncoder_GetEncodedResource(ReturnKey As ResourceKey, ReturnValue As Variant) As Boolean
    IResourceEncoder_GetEncodedResource = GetEncodedResource(ReturnKey, ReturnValue)
End Function

Private Sub IResourceEncoder_Reset()
    Call Reset
End Sub
