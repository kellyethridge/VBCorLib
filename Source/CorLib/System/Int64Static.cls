VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Int64Static"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: Int64Static
'

''
' Provides methods for manipulating a 64-bit signed value.
'
Option Explicit

''
' Represents a 64-bit signed integer.
'
' @param LowPart The least significant 32-bits of the integer value.
' @param HighPart The most significant 32-bits of the integer value.
Public Type Int64
    LowPart     As Long ' the ordering is important to remain consistant with memory layout of a 64-bit integer.
    HighPart    As Long
End Type


''
' Determines of two <a href="Int64Type.Int64.html">Int64</a> values are equal.
'
' @param Int64A The first value to compare.
' @param Int64B The second value to compare.
' @return <b>True</b> if the values are equal; otherwise, <b>False</b>.
Public Function Equals(ByRef Int64A As Int64, ByRef Int64B As Int64) As Boolean
    Equals = (Int64A.HighPart = Int64B.HighPart) And (Int64A.LowPart = Int64B.LowPart)
End Function

