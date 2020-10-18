VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringResourceEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: StringResourceEncoder
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

Private Const StringsPerBlock As Long = 16

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
' @remarks Any value can be passed in. No exception is thrown. If the value is not a string then the function returns False.
' <p>The <i>ResourceName</i> should be a numeric value.</p>
' <p>The LanguageID can be a <a href="CultureInfo.html">CultureInfo</a> object, a culture name, or an LCID. If
' the LanguageID is not supplied, then the current culture is used.</p>
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean
    Const MaxStringLength   As Long = 65535
    Const MinStringId       As Long = 1
    Const MaxStringId       As Long = 65535
    Dim Key     As ResourceKey
    Dim Block   As StringResourceBlock
    
    If CorVarType(Value) <> vbString Or Not IsInteger(ResourceName) Then
        Exit Function
    End If
    
    If Len(Value) > MaxStringLength Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", Environment.GetResourceString(Argument_MaxStringLength, MaxStringLength))
    If ResourceName < MinStringId Or ResourceName > MaxStringId Then _
        Throw Cor.NewArgumentOutOfRangeException("ResourceName", Environment.GetResourceString(ArgumentOutOfRange_Range, MinStringId, MaxStringId))
    
    ' Create our key to a specific resource entry into
    ' the .RES file. This is not the same as the StringID.
    ' However, the entry ID is derived from the String ID.
    Set Key = Cor.NewResourceKey(GetResourceEntryID(ResourceName), ResourceTypes.StringResource, GetLanguageID(LanguageID))
    Set Block = GetResourceBlock(Key)
    Block.AddString ResourceName, CStr(Value)
    
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
Public Function GetEncodedResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue() As Byte) As Boolean
    Dim Block As StringResourceBlock
    
    If mIterator Is Nothing Then
        Set mIterator = mStringResources.GetEnumerator
    End If
    
    If mIterator.MoveNext Then
        Set ReturnKey = mIterator.Key
        Set Block = mIterator.Value
        ReturnValue = Block.GetEncodedResource
        GetEncodedResource = True
    End If
End Function

''
' Releases all currently encoded values.
'
Public Sub Reset()
    mStringResources.Clear
    Set mIterator = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' String IDs are used 2 fold. The upper 12 bits are used to identify
' a specific resource entry in the .RES file. This entry contains
' 16 strings. The lower 4 bits are used to determine the index
' into the 16 strings of the specific resource entry.
Private Function GetResourceEntryID(ByVal StringId As Long) As Long
    ' shift the upper 12 bits right and add 1
    ' because we can never have a 0 ID. The 1
    ' will need to be subtracted when calculating
    ' back out the String ID.
    GetResourceEntryID = (StringId \ StringsPerBlock) + 1
End Function

''
' Get the resource entry based on the key. If the
' entry does not exist already, then create it.
'
Private Function GetResourceBlock(ByVal Key As ResourceKey) As StringResourceBlock
    If mStringResources.Contains(Key) Then
        Set GetResourceBlock = mStringResources(Key)
    Else
        Set GetResourceBlock = New StringResourceBlock
        mStringResources.Add Key, GetResourceBlock
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceEncoder
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResourceEncoder_Encode(Value As Variant, ResourceName As Variant, Optional ResourceType As Variant, Optional LanguageID As Variant) As Boolean
    IResourceEncoder_Encode = Encode(Value, ResourceName, ResourceType, LanguageID)
End Function

Private Function IResourceEncoder_GetEncodedResource(ReturnKey As ResourceKey, ReturnValue() As Byte) As Boolean
    IResourceEncoder_GetEncodedResource = GetEncodedResource(ReturnKey, ReturnValue)
End Function

Private Sub IResourceEncoder_Reset()
    Reset
End Sub
