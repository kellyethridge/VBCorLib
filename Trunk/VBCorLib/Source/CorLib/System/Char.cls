VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Char"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: Char
'
Option Explicit

''
' Indicates if a character is considered white space.
'
' @param CharOrString A string to check if character at specified index is white space.
' @param Index An index into the string to determine if a specific character is white space.
' This parameter is ignored if a non-string character is passed in.
' @return If the character is considered white space then True is returned, otherwise False is returned.
'
Public Function IsWhiteSpace(ByRef s As String, ByVal Index As Long) As Boolean
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
        
    Dim Ptr As Long
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsWhiteSpace = IsWhiteSpaceChar(MemWord(Ptr))
End Function

''
' Indicates if a character is considered white space.
'
' @param c An integer to check if is white space.
' @return If the character is considered white space then True is returned, otherwise False is returned.
'
Public Function IsWhiteSpaceChar(ByVal c As Integer) As Boolean
    Select Case c
        Case &H20, &HD, &H9, &HA, &HB, &HC, &H85, &HA0, &H1680, &H180E, _
             &H2000 To &H200A, _
             &H2028, &H2029, &H202F, &H205F, _
             &H3000
            IsWhiteSpaceChar = True
    End Select
End Function

Public Function IsHighSurrogate(ByRef s As String, ByVal Index As Long) As Boolean
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
    
    Dim Ptr As Long
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsHighSurrogate = IsHighSurrogateChar(MemWord(Ptr))
End Function

Public Function IsHighSurrogateChar(ByVal c As Long) As Boolean
    Select Case c
        Case &HD800& To &HDBFF&, &HFFFFD800 To &HFFFFDBFF
            IsHighSurrogateChar = True
    End Select
End Function

Public Function IsLowSurrogate(ByRef s As String, ByVal Index As Long) As Boolean
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
    
    Dim Ptr As Long
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsLowSurrogate = IsLowSurrogateChar(MemWord(Ptr))
End Function

Public Function IsLowSurrogateChar(ByVal c As Long) As Boolean
    Select Case c
        Case &HDC00& To &HDFFF&, &HFFFFDC00 To &HFFFFDFFF
            IsLowSurrogateChar = True
    End Select
End Function

Public Function IsSurrogate(ByRef s As String, ByVal Index As Long) As Boolean
    If IsHighSurrogate(s, Index) Then
        IsSurrogate = True
    ElseIf IsLowSurrogate(s, Index) Then
        IsSurrogate = True
    End If
End Function

Public Function IsSurrogateChar(ByVal c As Long) As Boolean
    If IsHighSurrogateChar(c) Then
        IsSurrogateChar = True
    ElseIf IsLowSurrogateChar(c) Then
        IsSurrogateChar = True
    End If
End Function
