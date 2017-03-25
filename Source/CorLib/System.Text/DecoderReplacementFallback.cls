VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecoderReplacementFallback"
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
' Module: DecoderReplacementFallback
'
Option Explicit
Implements IObject
Implements DecoderFallback

Private mDefaultString As String


Public Function CreateFallbackBuffer() As DecoderFallbackBuffer
    Set CreateFallbackBuffer = Cor.NewDecoderReplacementFallbackBuffer(Me)
End Function

Public Property Get DefaultString() As String
    DefaultString = mDefaultString
End Property

Public Property Get MaxCharCount() As Long
    MaxCharCount = Len(mDefaultString)
End Property

Public Function Equals(ByRef Value As Variant) As Boolean
    Dim Other As DecoderReplacementFallback
    
    On Error GoTo Catch
    Set Other = Value
    Equals = Other.DefaultString = mDefaultString
    
Catch:
End Function

Public Function GetHashCode() As Long
    GetHashCode = Object.GetHashCode(mDefaultString)
End Function

Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Replacement As String)
    Dim FoundHigh   As Boolean
    Dim i           As Long
    
    For i = 0 To Len(Replacement) - 1
        If Char.IsSurrogateStr(Replacement, i) Then
            If Char.IsHighSurrogateStr(Replacement, i) Then
                If FoundHigh Then
                    Exit For
                End If
                
                FoundHigh = True
            End If
            
            If Char.IsLowSurrogateStr(Replacement, i) Then
                If Not FoundHigh Then
                    FoundHigh = True
                    Exit For
                End If
                
                FoundHigh = False
            End If
        ElseIf FoundHigh Then
            Exit For
        End If
    Next

    If FoundHigh Then _
        Error.Argument Argument_InvalidCharSequenceNoIndex, "Replacement"

    mDefaultString = Replacement
End Sub

Private Sub Class_Initialize()
    Init "?"
End Sub


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
    IObject_GetHashCode = GetHashCode()
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString()
End Function
