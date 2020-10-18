VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IResourceDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: IResourceDecoder
'

''
' Provides an interface to allow a <b>ResourceReader</b> to decode
' resource types found within a .RES file.
'
Option Explicit

''
' Returns the type of resource this decoder can decode.
'
' @return Either the resource name or ID.
'
Public Function GetDecodeType() As Variant: End Function

''
' Decodes a raw resource data, returning the number of embedded resources created.
'
' @param Key The key associated with the raw resource data.
' @param Bytes The raw resource data to be decoded.
' @return The number of resources created from the raw data.
' @remarks The decoder needs to maintain the newly decoded resource. Only one
' resource is decoded at a time and then retrieved, so the need to maintain
' multiple decoded resources is unecessary.
'
Public Function Decode(ByVal Key As ResourceKey, ByRef Bytes() As Byte) As Long: End Function

''
' Returns the next resource key and data that was decoded from DecodeResource.
'
' @param ReturnKey The next resource key return to be returned.
' @param ReturnValue The next resource value to be returned.
' @param Returns True if the values were successfully set, otherwise False.
' @remarks Returns the next available decoded resource. Once that resource has
' been returned, it should not be returned again.
' <p>This function will be called as long as it returns True.</p>
'
Public Function GetResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue As Variant) As Boolean: End Function
