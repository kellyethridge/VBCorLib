VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitMapResourceEncoder"
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
' Module: BitmapResourceEncoder
'

''
' This encodes a multiple StdPicture bitmaps into individual arrays of bytes.
'
' @remarks In order to save a bitmap into a .RES file, it must be converted to an array of bytes so it can be written to the file.
' <p>Each time a bitmap is encoded, the array of bytes is stored internally for
' future retrieval. Multiple bitmaps can be encoded, each encoded bitmap will be
' stored with a <a href="ResourceKey.html">ResourceKey</a> for future retrieval.</p>
' @see ResourceWriter
'
Option Explicit
Implements IResourceEncoder

Private Const SIZEOF_BITMAPFILEHEADER   As Long = 14

Private mResources  As New Hashtable
Private mIterator   As IDictionaryEnumerator


''
' Encodes a bitmap into a <a href="ResourceKey.html">ResourceKey</a> and encoded byte array.
'
' @param Value The bitmap to be encoded.
' @param ResourceName The name or ID of the bitmap.
' @param ResourceType The type of value being encoded. This will always be ignored, and the value ResourceTypes.BitMapResource is used.
' @param LanguageID The LocaleID the bitmap is associated with.
' @return Returns <b>True</b> if the encoder was able to encode the value, <b>False</b> otherwise.
' @remarks Any value can be passed in. No exception is thrown. If the value is not a bitmap then the function returns <b>False</b>.
' <p>The <i>ResourceName</i> can be a String or a Numeric value.</p>
' <p>The LanguageID can be a <a href="CultureInfo.html">CultureInfo</a> object, a culture name, or an LCID. If the LanguageID is not supplied, then the current culture is used.</p>
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean
    Dim Pic             As IPicture
    Dim Stream          As IStream
    Dim BytesWritten    As Long
    Dim BytesToRead     As Long
    Dim Data()          As Byte
    Dim Key             As ResourceKey
    
    If Not IsPicture(Value) Then
        Exit Function
    End If
    
    Set Pic = Value
    
    ' Lets make sure the picture is actually a bitmap.
    If Pic.Type <> PICTYPE_BITMAP Then
        Exit Function
    End If
    
    ' Prepare to have the picture save itself out.
    CreateStreamOnHGlobal 0, BOOL_TRUE, Stream
    
    ' Save the picture, producing the bytes we need.
    Pic.SaveAsFile ByVal ObjPtr(Stream), True, BytesWritten
    
    ' Move to the first byte past the BITMAPFILEHEADER structure.
    ' 0.0014 as vbCurrency equals 14 as 64bit.
    Stream.Seek 0.0014@, STREAM_SEEK_SET, 0
    
    ' We don't include the size of the BITMAPFILEHEADER.
    BytesToRead = BytesWritten - SIZEOF_BITMAPFILEHEADER
    
    ' And read the bytes back out of the stream.
    ReDim Data(0 To BytesToRead - 1)
    Stream.Read Data(0), BytesToRead, 0
    
    Set Key = Cor.NewResourceKey(ResourceName, ResourceTypes.BitmapResource, GetLanguageID(LanguageID))
    mResources.Add Key, Data
    
    Encode = True
End Function

''
' Retrieves the next encoded resource in the encoder.
'
' @param OutKey This is set to a <a href="ResourceKey.html">ResourceKey</a> object that uniquely identify the encoded resource.
' @param OutValue This will be set to a byte array.
' @return Returns <b>True</b> if an encoded resource was returned, <b>False</b> otherwise.
' @remarks This function can be called multiple times to retrieve multiple resources.
' As long as thre is a resource to be returned, <b>True</b> will be returned, otherwise <b>False</b> will be returned.
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
    mResources.Clear
    Set mIterator = Nothing
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
