VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResStringTableEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridges
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
'    Module: ResStringTableEncoder
'

''
' Encodes strings into string tables.
'
' @remarks In a .RES file strings are stored into tables of up to
' 16 strings per table. Each table is uniquely identified based on
' the range of string ID's.
'
Option Explicit
Implements IResourceEncoder

Private Const STRINGS_PER_ENTRY As Long = 16


Private mStringResources    As New Hashtable
Private mIterator           As IDictionaryEnumerator



''
' Encodes a string into a string table.
'
' @param Value The string to be encoded.
' @param ResourceName The name of the string. This must be a numeric value.
' @param ResourceType The type of resource to be encoded.
' @param LanguageID The Locale ID this resource is associated with.
' @return Returns True if the encoder was able to encode the value, False otherwise.
' @remarks Any value can be passed in. No exception is thrown. If the value is not
' a string then the function returns False.
' <p>The <b>ResourceName</b> should be a Numeric value.</p>
' <p>The LanguageID can be a <b>CultureInfo</b> object, a culture name, or an LCID. If
' the LanguageID is not supplied, then the current culture is used.</p>
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean
    If IsMissing(ResourceType) Then
        ' If the type is missing, we can still handle the value
        ' if it is indeed a vbString datatype, otherwise, exit.
        If VarType(Value) <> vbString Then Exit Function
    Else
        ' The resource type must be the number 6, not a String "6".
        If Not IsInteger(ResourceType) Then Exit Function
        ' Check if it is a 6.
        If ResourceType <> ResourceTypes.StringResource Then Exit Function
    End If
    
    ' String resources require a numeric ID.
    If Not IsInteger(ResourceName) Then Exit Function
    
    ' Create our key to a specific resource entry into
    ' the .RES file. This is not the same as the StringID.
    ' However, the entry ID is derived from the String ID.
    Dim Key As ResourceKey
    Set Key = Cor.NewResourceKey(GetResourceEntryID(ResourceName), ResourceTypes.StringResource, GetLanguageID(LanguageID))
    
    Dim ResourceEntry As ResStringResource
    Set ResourceEntry = GetResourceEntry(Key)
    
    Call ResourceEntry.AddString(ResourceName, Value)
    
    Encode = True
End Function

''
' Returns the next resource entry byte array.
'
' @param ReturnKey Our own key for the resource entry.
' @param ReturnValue The next byte array of a ResStringResource.
' @return If we have returned something, Return True, otherwise return False.
' @remarks Each resource can contain up to 16 of the strings previously encoded.
'
Public Function GetEncodedResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue As Variant) As Boolean
    ' Prepare our iterator. We only create it once.
    ' If we reset, then we'll create it again.
    If mIterator Is Nothing Then
        Set mIterator = mStringResources.GetEnumerator
    End If
    
    ' Attempt to move to the next ResStringResource.
    If mIterator.MoveNext Then
        ' The key is the same one we created in the Encode method.
        ' It is already set up to what we want.
        Set ReturnKey = mIterator.Key
        
        ' We need to cast our value to a ResStringResource
        ' so we can easily access the functions.
        Dim ResourceEntry As ResStringResource
        Set ResourceEntry = mIterator.Value
        
        ' Returns a byte array of the resource entry.
        ReturnValue = ResourceEntry.GetEncodedResource
        
        ' Signal that we have indeed set return values.
        GetEncodedResource = True
    End If
End Function

''
' Releases all currently encoded values.
'
Public Sub Reset()
    Call mStringResources.Clear
    Set mIterator = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' String IDs are used 2 fold. The upper 12 bits are used to identify
' a specific resource entry in the .RES file. This entry contains
' 16 strings. The lower 4 bits are used to determine the index
' into the 16 strings of the specific resource entry.
Private Function GetResourceEntryID(ByVal StringID As Long)
    If StringID < 1 Or StringID > 65535 Then _
        Throw Cor.NewArgumentOutOfRangeException("ResourceName", "String ID's can only be 1 to 65535.")
    
    ' shift the upper 12 bits right and add 1
    ' because we can never have a 0 ID. The 1
    ' will need to be subtracted when calculating
    ' back out the String ID.
    GetResourceEntryID = (StringID \ STRINGS_PER_ENTRY) + 1
End Function

''
' Get the resource entry based on the key. If the
' entry does not exist already, then create it.
'
Private Function GetResourceEntry(ByVal Key As ResourceKey) As ResStringResource
    If mStringResources.Contains(Key) Then
        Set GetResourceEntry = mStringResources(Key)
    Else
        Set GetResourceEntry = New ResStringResource
        Call mStringResources.Add(Key, GetResourceEntry)
    End If
End Function


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
