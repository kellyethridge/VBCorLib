VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderExceptionFallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: EncoderExceptionFallback
'
Option Explicit
Implements IObject
Implements EncoderFallback


''
' Returns an encoder fallback buffer that throws an exception if it cannot convert a character sequence to a byte sequence.
'
' @return An encoder fallback buffer that throws an exception when it cannot encode a character sequence.
'
Public Function CreateFallbackBuffer() As EncoderFallbackBuffer
    Set CreateFallbackBuffer = New EncoderExceptionFallbackBuffer
End Function

''
' Gets the maximum number of characters this instance can return.
'
' @return This always returns zero.
'
Public Property Get MaxCharCount() As Long
    MaxCharCount = 0
End Property

''
' Determines whether the specified value passed in is equal to the current object.
'
' @param Value The value to compare to the current object.
' @return <b>True</b> if the specified value is equal to the current object; otherwise, <b>False</b>.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Serves as the default hash function.
'
' @return A hash code for the current object.
' @see IObject
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string that represents the current object.
'
' @return A string the represents the current object.
' @see IObject
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   EncoderFallback
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EncoderFallback_CreateFallbackBuffer() As EncoderFallbackBuffer
    Set EncoderFallback_CreateFallbackBuffer = CreateFallbackBuffer
End Function

Private Function EncoderFallback_Equals(Value As Variant) As Boolean
    EncoderFallback_Equals = Equals(Value)
End Function

Private Function EncoderFallback_GetHashCode() As Long
    EncoderFallback_GetHashCode = GetHashCode
End Function

Private Property Get EncoderFallback_MaxCharCount() As Long
    EncoderFallback_MaxCharCount = MaxCharCount
End Property

Private Function EncoderFallback_ToString() As String
    EncoderFallback_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
