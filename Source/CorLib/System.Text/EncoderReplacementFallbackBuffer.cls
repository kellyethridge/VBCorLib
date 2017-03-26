VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderReplacementFallbackBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: EncoderReplacementFallbackBuffer
'

''
' Represents a substitute input string that is used when the original input character cannot be encoded. This class cannot be inherited.
'
Option Explicit
Implements IObject
Implements EncoderFallbackBuffer

Private mDefaultString()    As Integer
Private mFallbackCount      As Long
Private mFallbackIndex      As Long


Public Function Fallback(ByVal UnknownCharOrHighSurrogate As Long, Optional ByRef UnknownLowSurrogate As Variant) As Boolean
    If IsMissing(UnknownLowSurrogate) Then
        Fallback = UnknownCharFallback(UnknownCharOrHighSurrogate)
    Else
        Fallback = UnknownSurrogatePairFallback(UnknownCharOrHighSurrogate, UnknownLowSurrogate)
    End If
End Function

Public Function GetNextChar() As Integer
    mFallbackIndex = mFallbackIndex + 1
    mFallbackCount = mFallbackCount - 1
    
    If mFallbackCount >= 0 Then
        GetNextChar = mDefaultString(mFallbackIndex)
    End If
End Function

Public Function MovePrevious() As Boolean
    If mFallbackIndex >= 0 Then
        mFallbackCount = mFallbackCount + 1
        mFallbackIndex = mFallbackIndex - 1
        MovePrevious = True
    End If
End Function

Public Property Get Remaining() As Long
    If mFallbackCount > 0 Then
        Remaining = mFallbackCount
    End If
End Property

Public Sub Reset()
    mFallbackCount = -1
    mFallbackIndex = -1
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
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Fallback As EncoderReplacementFallback)
    mDefaultString = CorString.ToCharArray(Fallback.DefaultString & Fallback.DefaultString)
End Sub

Private Sub Class_Initialize()
    Reset
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function UnknownCharFallback(ByVal UnknownChar As Long) As Boolean
    UnknownChar = UnknownChar And &HFFFF&
    
    If mFallbackCount > 0 Then
        Dim Ch As Long
        Ch = UnknownChar
        
        If IsHighSurrogate(UnknownChar) And mFallbackCount >= 0 Then
            If IsLowSurrogate(mDefaultString(mFallbackIndex + 1)) Then
                Ch = ConvertToUtf32(UnknownChar, mDefaultString(mFallbackIndex + 1))
            End If
        End If
        
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_RecursiveFallback, Ch), "Chars")
    End If
    
    mFallbackIndex = -1
    mFallbackCount = Len1D(mDefaultString) \ 2
    UnknownCharFallback = mFallbackCount > 0
End Function

Private Function UnknownSurrogatePairFallback(ByVal UnknownHighSurrogate As Long, ByRef UnknownLowSurrogate As Variant) As Boolean
    Dim High    As Long
    Dim Low     As Long
    
    High = UnknownHighSurrogate And &HFFFF&
    Low = CLng(UnknownLowSurrogate) And &HFFFF&
    
    If Not IsHighSurrogate(High) Then _
        Throw Cor.NewArgumentOutOfRangeException("UnknownCharOrHighSurrogate", Environment.GetResourceString(ArgumentOutOfRange_Range, &HD800&, &HDBFF&))
    If Not IsLowSurrogate(Low) Then _
        Throw Cor.NewArgumentOutOfRangeException("UnknownLowSurrogate", Environment.GetResourceString(ArgumentOutOfRange_Range, &HDC00&, &HDFFF&))
    If mFallbackCount > 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_RecursiveFallback, ConvertToUtf32(High, Low)), "Chars")
    
    mFallbackIndex = -1
    mFallbackCount = Len1D(mDefaultString)
    UnknownSurrogatePairFallback = mFallbackCount > 0
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   EncoderFallbackBuffer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EncoderFallbackBuffer_Equals(Value As Variant) As Boolean
    EncoderFallbackBuffer_Equals = Equals(Value)
End Function

Private Function EncoderFallbackBuffer_Fallback(ByVal UnknownCharOrHighSurrogate As Long, Optional LowSurrogate As Variant) As Boolean
    EncoderFallbackBuffer_Fallback = Fallback(UnknownCharOrHighSurrogate, LowSurrogate)
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
