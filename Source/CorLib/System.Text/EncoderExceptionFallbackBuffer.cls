VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderExceptionFallbackBuffer"
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
' Module: EncoderExceptionFallbackBuffer
'
Option Explicit
Implements IObject
Implements EncoderFallbackBuffer


Public Function Fallback(ByVal UnknownCharOrHighSurrogate As Long, Optional ByRef UnknownLowSurrogate As Variant, Optional ByVal Index As Long) As Boolean
    Dim UTF32 As Long
    
    If Not IsMissing(UnknownLowSurrogate) Then
        If Not Char.IsHighSurrogate(UnknownCharOrHighSurrogate) Then _
            Throw Cor.NewArgumentOutOfRangeException("UnknownCharOrHighSurrogate", Environment.GetResourceString(ArgumentOutOfRange_Range, &HD800, &HDBFF))
        If Not Char.IsLowSurrogate(UnknownLowSurrogate) Then _
            Throw Cor.NewArgumentOutOfRangeException("UnknownLowSurrogate", Environment.GetResourceString(ArgumentOutOfRange_Range, &HDC00, &HDFFF))
        
        UTF32 = Char.ConvertToUtf32(UnknownCharOrHighSurrogate, UnknownLowSurrogate)
        
        Throw Cor.NewEncoderFallbackException3( _
            Environment.GetResourceString(Argument_InvalidCodePageConversionIndex, UTF32, Index), _
            UnknownCharOrHighSurrogate, UnknownLowSurrogate, Index)
    End If
    
    Throw Cor.NewEncoderFallbackException2( _
        Environment.GetResourceString(Argument_InvalidCodePageConversionIndex, UnknownCharOrHighSurrogate, Index), _
        UnknownCharOrHighSurrogate, Index)
End Function

Public Function GetNextChar() As Integer
    GetNextChar = 0
End Function

Public Function MovePrevious() As Boolean
    MovePrevious = False
End Function

Public Property Get Remaining() As Long
    Remaining = 0
End Property

Public Sub Reset()
    ' do nothing
End Sub


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
'   EncoderFallbackBuffer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EncoderFallbackBuffer_Equals(Value As Variant) As Boolean
    EncoderFallbackBuffer_Equals = Equals(Value)
End Function

Private Function EncoderFallbackBuffer_Fallback(ByVal UnknownCharOrHighSurrogate As Long, Optional UnknownLowSurrogate As Variant, Optional ByVal Index As Long) As Boolean
    EncoderFallbackBuffer_Fallback = Fallback(UnknownCharOrHighSurrogate, UnknownLowSurrogate, Index)
End Function

Private Function EncoderFallbackBuffer_GetHashCode() As Long
    EncoderFallbackBuffer_GetHashCode = GetHashCode
End Function

Private Function EncoderFallbackBuffer_GetNextChar() As Integer
    EncoderFallbackBuffer_GetNextChar = GetNextChar
End Function

Private Function EncoderFallbackBuffer_MovePrevious() As Boolean
    EncoderFallbackBuffer_MovePrevious = MovePrevious
End Function

Private Property Get EncoderFallbackBuffer_Remaining() As Long
    EncoderFallbackBuffer_Remaining = Remaining
End Property

Private Sub EncoderFallbackBuffer_Reset()
    Reset
End Sub

Private Function EncoderFallbackBuffer_ToString() As String
    EncoderFallbackBuffer_ToString = ToString
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
