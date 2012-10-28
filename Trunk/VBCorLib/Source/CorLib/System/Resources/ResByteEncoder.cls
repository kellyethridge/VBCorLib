VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResByteEncoder"
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
'    Module: ResByteEncoder
'

''
' This will accept a byte array to be encoded into a .RES file.
'
' @remarks Use this encoder to create a <b>ResourceKey</b> and byte array
' to be included in a .RES file.
' <p>Multiple resources can be encoded. Each encoded resource will be
' stored internally for later retrieval.</p>
' @see ResourceWriter
'
Option Explicit
Implements IResourceEncoder

Private mResources  As New Hashtable
Private mIterator   As IDictionaryEnumerator



''
' The byte array is assumed to be the encoded data and is written
' to a .RES file as is.
'
' @param Value The byte array to be written to the .RES file.
' @param ResourceName The name or ID of the bitmap.
' @param ResourceType The type of value being encoded.
' @param LanguageID The LocaleID the bitmap is associated with.
' @return Returns True if the encoder was able to encode the value, False otherwise.
' @remarks Any value can be passed in. No exception is thrown. If the value is not
' a bitmap then the function returns False.
' <p>The <b>ResourceName</b> can be a String or a Numeric value.</p>
' <p>The LanguageID can be a <b>CultureInfo</b> object, a culture name, or an LCID. If
' the LanguageID is not supplied, then the current culture is used.</p>
' <p>Null arrays are not encoded. False will be returned.</p>
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean
    If VarType(Value) <> vbByteArray Then Exit Function
    If cArray.IsNull(Value) Then Exit Function
    
    Dim Key As ResourceKey
    Set Key = Cor.NewResourceKey(ResourceName, ResourceType, GetLanguageID(LanguageID))
    Call mResources.Add(Key, Value)
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
    Set mIterator = Nothing
    Set mResources = New Hashtable
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
