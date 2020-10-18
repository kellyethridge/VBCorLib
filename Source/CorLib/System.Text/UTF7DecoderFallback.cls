VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF7DecoderFallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
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
' Module: UTF7DecoderFallback
'
Option Explicit
Implements IObject
Implements DecoderFallback


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DecoderFallback
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DecoderFallback_CreateFallbackBuffer() As DecoderFallbackBuffer
    Set DecoderFallback_CreateFallbackBuffer = New UTF7DecoderFallbackBuffer
End Function

Private Function DecoderFallback_Equals(Value As Variant) As Boolean
    DecoderFallback_Equals = IObject_Equals(Value)
End Function

Private Function DecoderFallback_GetHashCode() As Long
    DecoderFallback_GetHashCode = IObject_GetHashCode
End Function

Private Property Get DecoderFallback_MaxCharCount() As Long
    DecoderFallback_MaxCharCount = 1
End Property

Private Function DecoderFallback_ToString() As String
    DecoderFallback_ToString = IObject_ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    On Error GoTo Catch
    IObject_Equals = TypeOf Value Is UTF7DecoderFallback
Catch:
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = 984
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function
