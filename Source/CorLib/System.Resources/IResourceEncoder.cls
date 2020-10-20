VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IResourceEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: IResourceEncoder
'

''
' Provides an interface for classes to be used to encode data
' into byte arrays used by the <a href="ResourceWriter.html">ResourceWriter</a>.
'
' @remarks An encoder needs to hold multiple encoded resources. Once all resources
' have been encoded, each resource will be requested one at a time. The same
' resource should never be returned more than once.
'
'@Folder("CorLib.System.Resources")
Option Explicit

''
' Requests that the encoder release all currently encoded values.
'
Public Sub Reset(): End Sub

''
' Encodes a value into a byte array for later retrieval.
'
' @param Value The value to be encoded into a resource compatible byte array.
' @param ResourceName The identifier for the specific resource. Should be converted to a string or numeric value.
' @param ResourceType The type of resource. This can be a string or numeric value. If the is value is not supplied, then an attempt to determine the type should be made based on the Value itself.
' @param LanguageID The Locale identifier this resource belongs to. If this value is not supplied, then the current culture's LCID should be used.
' @return Returns <b>True</b> if the encoder accepted and encoded the value, <b>False</b> otherwise.
' @remarks This function should not throw an exception. If the encoder cannot handle the the value for encoding, it should just return <b>False</b>.
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean: End Function

''
' Returns the next available resource from the encoder.
'
' @param OutKey This is the unique identifier for the resource to be returned.
' @param OutValue A byte array representing the encoded resource value.
' @return Returns <b>True</b> if a resource has been returned, <b>False</b> otherwise.
' @remarks This function should not throw an exception. If there are no more resources to be returned, then False should just be returned.
'
Public Function GetEncodedResource(ByRef OutKey As ResourceKey, ByRef OutValue() As Byte) As Boolean: End Function
