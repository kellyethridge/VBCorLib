Attribute VB_Name = "CharMethods"
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
' Module: CharMethods
'

''
' This module was created to optimize method calls. This library doesn't use the Char
' static class method because these are types of methods that will likely be used in looping scenarios.
' The public facing Char class forwards calls to this module for clients of this library.
'
' These methods do no validation. It is assumed the caller validated the arguments.
Option Explicit

Public Const UnicodePlane1Start As Long = &H10000
Public Const UnicodePlane16End  As Long = &H10FFFF
Public Const SurrogateStart     As Long = &HD800&
Public Const SurrogateEnd       As Long = &HDFFF&
Public Const HighSurrogateStart As Long = &HD800&
Public Const HighSurrogateEnd   As Long = &HDBFF&
Public Const LowSurrogateStart  As Long = &HDC00&
Public Const LowSurrogateEnd    As Long = &HDFFF&
Public Const vbSizeOfUTF32Char  As Long = 4
Public Const MaxCharValue       As Long = 65535
Public Const MinCharValue       As Long = -32768


Public Function IsValidChar(ByVal Value As Long) As Boolean
    Select Case Value
        Case MinCharValue To MaxCharValue
            IsValidChar = True
    End Select
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

Public Function IsHighSurrogate(ByVal c As Long) As Boolean
    Select Case c And &HFFFF&
        Case HighSurrogateStart To HighSurrogateEnd
            IsHighSurrogate = True
    End Select
End Function

Public Function IsLowSurrogate(ByVal c As Long) As Boolean
    Select Case c And &HFFFF&
        Case LowSurrogateStart To LowSurrogateEnd
            IsLowSurrogate = True
    End Select
End Function

Public Function IsSurrogate(ByVal c As Long) As Boolean
    Select Case c And &HFFFF&
        Case SurrogateStart To SurrogateEnd
            IsSurrogate = True
    End Select
End Function

Public Function ConvertToUtf32(ByVal HighSurrogate As Long, ByVal LowSurrogate As Long) As Long
    ConvertToUtf32 = (HighSurrogate And &H3FF) * vbShift10Bits + (LowSurrogate And &H3FF) + UnicodePlane1Start
End Function

