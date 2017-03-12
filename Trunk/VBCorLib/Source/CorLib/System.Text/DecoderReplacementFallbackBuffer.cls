VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecoderReplacementFallbackBuffer"
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
' Module: DecoderReplacementFallbackBuffer
'
Option Explicit
Implements IObject
Implements DecoderFallbackBuffer

Private mFallbackChars()    As Integer
Private mFallbackLength     As Long
Private mFallbackCount      As Long
Private mFallbackIndex      As Long


Public Function Fallback(ByRef BytesUnknown() As Byte, ByVal Index As Long) As Boolean
    If mFallbackCount > 0 Then _
        ThrowLastBytesRecursive BytesUnknown

    If mFallbackLength > 0 Then
        mFallbackCount = mFallbackLength
        mFallbackIndex = -1
        Fallback = True
    End If
End Function

Private Sub ThrowLastBytesRecursive(ByRef BytesUnknown() As Byte)
    Dim sb  As StringBuilder
    Dim i   As Long
    
    Set sb = StringBuilderCache.Acquire
    
    For i = LBound(BytesUnknown) To CorMath.Min(LBound(BytesUnknown) + 20, UBound(BytesUnknown))
        If sb.Length > 0 Then
            sb.AppendChar vbSpaceChar
        End If
        
        sb.AppendFormat "x{0:X2}", BytesUnknown(i)
    Next
    
    If i <= UBound(BytesUnknown) Then
        sb.Append " ..."
    End If
    
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_RecursiveFallbackBytes, StringBuilderCache.GetStringAndRelease(sb)), "BytesUnknown")
End Sub

Public Function GetNextChar() As Integer
    mFallbackCount = mFallbackCount - 1
    mFallbackIndex = mFallbackIndex + 1
    
    ' the compiled version will have overflow checks off
    ' so we check here and let it continue.
    If mFallbackCount = &H7FFFFFFF Then
        mFallbackCount = -1
    End If
    
    If mFallbackCount >= 0 Then
        GetNextChar = mFallbackChars(mFallbackIndex)
    End If
End Function

Public Function MovePrevious() As Boolean
    If mFallbackCount >= -1 And mFallbackIndex >= 0 Then
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

Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Fallback As DecoderReplacementFallback)
    mFallbackChars = CorString.ToCharArray(Fallback.DefaultString)
    mFallbackLength = Len(Fallback.DefaultString)
End Sub

Private Sub Class_Initialize()
    Reset
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DecoderFallbackBuffer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DecoderFallbackBuffer_Fallback(BytesUnknown() As Byte, ByVal Index As Long) As Boolean
    DecoderFallbackBuffer_Fallback = Fallback(BytesUnknown, Index)
End Function

Private Function DecoderFallbackBuffer_GetNextChar() As Integer
    DecoderFallbackBuffer_GetNextChar = GetNextChar
End Function

Private Function DecoderFallbackBuffer_MovePrevious() As Boolean
    DecoderFallbackBuffer_MovePrevious = MovePrevious
End Function

Private Property Get DecoderFallbackBuffer_Remaining() As Long
    DecoderFallbackBuffer_Remaining = Remaining
End Property

Private Sub DecoderFallbackBuffer_Reset()
    Reset
End Sub


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
