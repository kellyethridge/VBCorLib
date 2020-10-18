VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CursorResourceEncoder"
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
' Module: CursorResourceEncoder
'

''
' Encodes a StdPicture Cursor to byte data used in .RES files.
'
' @see ResourceWriter
'
Option Explicit
Implements IResourceEncoder

Private mResources  As New Hashtable
Private mIterator   As IDictionaryEnumerator


''
' Encodes the StdPicture cursor.
'
' @param Value The StdPicture to be encoded.
' @param ResourceName The identifier for the resource value. This can be a String or Number.
' @param ResourceType Identifies the type of resource this is. If this is not supplied, then
' the type is derived from the value.
' @param LanguageID The LocaleID the value is associated with.
' @return Returns True if the encoder was able to encode the value, False otherwise.
' @remarks Any value can be passed in. No exception is thrown. If the value is not a cursor then the function returns False.
' <p>The <i>ResourceName</i> can be a String or a Numeric value.</p>
' <p>The LanguageID can be a <a href="CultureInfo.html">CultureInfo</a> object, a culture name, or an LCID. If
' the LanguageID is not supplied, then the current culture is used.</p>
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean
    Const HeaderSizeMinusHotSpotSize As Long = 22 - 4
    Dim XHotSpot    As Long
    Dim YHotSpot    As Long
    Dim Handle      As Long
    Dim Bytes()     As Byte
    
    If IsPicture(Value) Then
        Dim Pic As IPicture
        Set Pic = Value
        Handle = Pic.Handle
    ElseIf IsInteger(Value) Then
        Handle = Value
    Else
        Exit Function
    End If
    
    Bytes = CursorResourceFromHandle(Handle)
    
    If CorArray.IsNull(Bytes) Then
        Exit Function
    End If
        
    XHotSpot = Bytes(10)
    YHotSpot = Bytes(12)
    Buffer.BlockCopy Bytes, HeaderSizeMinusHotSpotSize, Bytes, 0, CorArray.Length(Bytes) - HeaderSizeMinusHotSpotSize
    ReDim Preserve Bytes(0 To UBound(Bytes) - HeaderSizeMinusHotSpotSize)
    Bytes(0) = XHotSpot
    Bytes(2) = YHotSpot
    
    Dim Key As ResourceKey
    Set Key = Cor.NewResourceKey(ResourceName, ResourceTypes.CursorResource, GetLanguageID(LanguageID))

    mResources.Add Key, Bytes
    
    Encode = True
End Function

''
' Retrieves the next encoded resource in the encoder.
'
' @param ReturnKey This is set to a <a href="ResourceKey.html">ResourceKey</a> object that uniquely identify the encoded resource.
' @param ReturnValue This will be set to a byte array.
' @return Returns True if an encoded resource was returned, False otherwise.
' @remarks This function can be called multiple times to retrieve multiple resources.
' As long as a resource is returned, True will be returned, otherwise False will be returned.
'
Public Function GetEncodedResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue() As Byte) As Boolean
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
    mResources.Clear
    Set mIterator = Nothing
End Sub


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
