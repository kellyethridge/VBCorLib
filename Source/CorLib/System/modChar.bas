Attribute VB_Name = "modChar"
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
' Module: modChar
'

''
' This module was created to optimize method calls. This library doesn't use the Char
' static class method because these are types of methods that will likely be used in looping scenarios.
' The public facing Char class forwards calls to this module for clients of this library.
'
Option Explicit

Public Function Compare(ByVal a As Long, ByVal b As Long) As Long
    a = a And &HFFFF&
    b = b And &HFFFF&
    
    If a < b Then
        Compare = -1
    ElseIf a > b Then
        Compare = 1
    End If
End Function

Public Function Equals(ByVal a As Long, ByVal b As Long) As Boolean
    Equals = (a And &HFFFF&) = (b And &HFFFF&)
End Function

Public Function IsWhiteSpaceStr(ByRef s As String, ByVal Index As Long) As Boolean
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
        
    Dim Ptr As Long
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsWhiteSpaceStr = IsWhiteSpace(MemWord(Ptr))
End Function

Public Function IsWhiteSpace(ByVal c As Long) As Boolean
    Select Case c
        Case &H20, &HD, &H9, &HA, &HB, &HC, &H85, &HA0, &H1680, &H180E, _
             &H2000 To &H200A, _
             &H2028, &H2029, &H202F, &H205F, _
             &H3000
            IsWhiteSpace = True
    End Select
End Function

Public Function IsHighSurrogateStr(ByRef s As String, ByVal Index As Long) As Boolean
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
    
    Dim Ptr As Long
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsHighSurrogateStr = IsHighSurrogate(MemWord(Ptr))
End Function

Public Function IsHighSurrogate(ByVal c As Long) As Boolean
    Select Case c
        Case &HD800& To &HDBFF&, &HFFFFD800 To &HFFFFDBFF
            IsHighSurrogate = True
    End Select
End Function

Public Function IsLowSurrogateStr(ByRef s As String, ByVal Index As Long) As Boolean
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
    
    Dim Ptr As Long
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsLowSurrogateStr = IsLowSurrogate(MemWord(Ptr))
End Function

Public Function IsLowSurrogate(ByVal c As Long) As Boolean
    Select Case c
        Case &HDC00& To &HDFFF&, &HFFFFDC00 To &HFFFFDFFF
            IsLowSurrogate = True
    End Select
End Function

Public Function IsSurrogateStr(ByRef s As String, ByVal Index As Long) As Boolean
    If IsHighSurrogateStr(s, Index) Then
        IsSurrogateStr = True
    ElseIf IsLowSurrogateStr(s, Index) Then
        IsSurrogateStr = True
    End If
End Function

Public Function IsSurrogate(ByVal c As Long) As Boolean
    If IsHighSurrogate(c) Then
        IsSurrogate = True
    ElseIf IsLowSurrogate(c) Then
        IsSurrogate = True
    End If
End Function

