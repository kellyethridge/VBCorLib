VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitArrayStatic"
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
' Module: BitArrayStatic
'

''
' Provides creation methods from various types of arrays.
'
' @remarks This class cannot be directly instantiated. In order to
' use the methods in the class, use the <b>BitArray.*</b> syntax.
' <pre>
' Set Bits = BitArray.FromLongs(LongArray)
' </pre>
' @see BitArray
'
Option Explicit

''
' Returns a BitArray object created from an array of Longs. Each bit
' in each long represents an associated bit in the bit array.
'
' @param values An array of Long datatypes.
' @return A BitArray initialized to the number of bits an their
' state in the values array.
'
Public Function FromLongs(ByRef Values() As Long) As BitArray
    Set FromLongs = New BitArray
    FromLongs.InitFromLongs Values
End Function

''
' Returns a BitArray object created from an array of Bytes. Each bit
' in each byte represents an associated bit in the bit array.
'
' @param values An array of Byte datatypes.
' @return A BitArray initialized to the number of bits and their
' state in the values array.
'
Public Function FromBytes(ByRef Values() As Byte) As BitArray
    Set FromBytes = New BitArray
    FromBytes.InitFromBytes Values
End Function

''
' Returns a BitArray object created from an array of Booleans. Each boolean
' represents a single bit in the bit array.
'
' @param values An array of Boolean datatypes.
' @return A BitArray initialized to the number of bits as there are
' boolean elements in the array. Each bit is set to the corrisponding
' boolean value in the array.
'
Public Function FromBooleans(ByRef Values() As Boolean) As BitArray
    Set FromBooleans = New BitArray
    FromBooleans.InitFromBooleans Values
End Function
