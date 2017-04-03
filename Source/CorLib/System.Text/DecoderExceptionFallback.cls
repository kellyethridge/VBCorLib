VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecoderExceptionFallback"
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
' Module: DecoderExceptionFallback
'

''
' Provides a failure-handling mechanism, called a fallback, for an encoded input byte sequence that cannot be converted to an input character. The fallback throws an exception instead of decoding the input byte sequence.
'
Option Explicit
Implements IObject
Implements DecoderFallback


''
' Initializes a new instance of the <a href="DecoderExceptionFallbackBuffer.html">DecoderExceptionFallbackBuffer</a> class.
'
Public Function CreateFallbackBuffer() As DecoderFallbackBuffer
    Set CreateFallbackBuffer = New DecoderExceptionFallbackBuffer
End Function

''
' Return the maximum number of characters the current <a href="DecoderFallback.html">DecoderFallback</a> object can return.
'
' @return The maximum number of characters the current <a href="DecoderFallback.html">DecoderFallback</a> object can return.
' @remarks For this <a href="DecoderFallback.html">DecoderFallback</a> this property will aways return zero.
'
Public Property Get MaxCharCount() As Long
    MaxCharCount = 0
End Property

''
' Determines equality with another value.
'
' @param Value The other value to check equality against.
' @return <b>True</b> if both values are considered equal; otherwise <b>False</b>.
' @remarks A <a href="DecoderExceptionFallback.html">DecoderExceptionFallback</a> is considered
' equal to another object if it is also a <a href="DecoderExceptionFallback.html">DecoderExceptionFallback</a> object.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Dim Other As DecoderExceptionFallback
    
    On Error GoTo Catch
    Set Other = Value
    Equals = Not Other Is Nothing
    
Catch:
End Function

''
' Gets a partially unique value that identifies this object.
'
' @return A number that is generally unique to this object.
' @remarks <p>The <a href="DecoderExceptionFallback.html">DecoderExceptionFallback</a> does not have
' any internal state that would normally be used to identify this instance. So, every instance of the
' <a href="DecoderExceptionFallback.html">DecoderExceptionFallback</a> class will return the same
' hash code.</p>
'
Public Function GetHashCode() As Long
    GetHashCode = 879
End Function

''
' Gets a string representation of this object instance.
'
' @return A string describing this object instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DecoderFallback
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DecoderFallback_CreateFallbackBuffer() As DecoderFallbackBuffer
    Set DecoderFallback_CreateFallbackBuffer = CreateFallbackBuffer
End Function

Private Function DecoderFallback_Equals(Value As Variant) As Boolean
    DecoderFallback_Equals = Equals(Value)
End Function

Private Function DecoderFallback_GetHashCode() As Long
    DecoderFallback_GetHashCode = GetHashCode
End Function

Private Property Get DecoderFallback_MaxCharCount() As Long
    DecoderFallback_MaxCharCount = MaxCharCount
End Property

Private Function DecoderFallback_ToString() As String
    DecoderFallback_ToString = ToString
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
