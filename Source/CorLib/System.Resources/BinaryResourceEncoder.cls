VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryResourceEncoder"
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
' Module: BinaryResourceEncoder
'

''
' This will accept a byte array to be encoded into a .RES file.
'
' @remarks Multiple resources can be encoded. Each encoded resource will be stored internally for later retrieval.
' @see ResourceWriter
'
Option Explicit
Implements IResourceEncoder

Private mResources  As New Hashtable
Private mIterator   As IDictionaryEnumerator


''
' Stores a byte array to be written out to a resource (.RES) file.
'
' @param Value The byte array to be written to the .RES file.
' @param ResourceName The name or ID of the resource.
' @param ResourceType The type of value being encoded.
' @param LanguageID The LocaleID the bitmap is associated with.
' @return Returns True if the encoder was able to encode the value, False otherwise.
' @remarks Any value can be passed in. No exception is thrown. If the value is not a Byte array then the function returns False.
' <p>The <i>ResourceName</i> and <i>ResourceType</i> must be a String or a Numeric value.</p>
' <p>The LanguageID can be a <a href="CultureInfo.html">CultureInfo</a> object, a culture name, or an LCID. If the LanguageID is not supplied, then the current culture is used.</p>
' <p>Null arrays are not encoded. False will be returned.</p>
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean
    ValidateResourceName ResourceName
    
    If VarType(Value) <> vbByteArray Then
        Exit Function
    End If
    
    If CorArray.IsNull(Value) Then
        Exit Function
    End If
    
    Dim Key As ResourceKey
    Set Key = Cor.NewResourceKey(ResourceName, CVarOrDefault(ResourceType, "CUSTOM"), GetLanguageID(LanguageID))
    mResources.Add Key, Value
    Encode = True
End Function



''
' Retrieves the next encoded resource in the encoder.
'
' @param OutKey This is set to a <a href="ResourceKey.html">ResourceKey</a> object that uniquely identify the encoded resource.
' @param OutValue This will be set to a byte array.
' @return Returns <b>True</b> if an encoded resource was returned, <b>False</b> otherwise.
' @remarks This function can be called multiple times to retrieve multiple resources.
' As long as a resource is returned, <b>True</b> will be returned, otherwise <b>False</b> will be returned.
'
Public Function GetEncodedResource(ByRef OutKey As ResourceKey, ByRef OutValue() As Byte) As Boolean
    If mIterator Is Nothing Then
        Set mIterator = mResources.GetEnumerator
    End If
    
    If mIterator.MoveNext Then
        Set OutKey = mIterator.Key
        OutValue = mIterator.Value
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
'   IResourceEncoder
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResourceEncoder_Encode(Value As Variant, ResourceName As Variant, Optional ResourceType As Variant, Optional LanguageID As Variant) As Boolean
    IResourceEncoder_Encode = Encode(Value, ResourceName, ResourceType, LanguageID)
End Function

Private Function IResourceEncoder_GetEncodedResource(OutKey As ResourceKey, OutValue() As Byte) As Boolean
    IResourceEncoder_GetEncodedResource = GetEncodedResource(OutKey, OutValue)
End Function

Private Sub IResourceEncoder_Reset()
    Reset
End Sub
