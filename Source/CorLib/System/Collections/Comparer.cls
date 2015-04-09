VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: Comparer
'

''
' Provides a default mechanism to compare two values. The values must be of the
' same datatype to be compared. Otherwise an exception will be thrown.
'
' @remarks ' <p>The default comparer requires that the two datatypes being compared be
' the same. This includes comparing two numbers such that comparing a vbLong type to
' a vbInteger type will result in an exception.<br>
' Strings will not be converted to numbers for comparison, even though VB supports this
' action. The comparisons of strings is case-sensitive. For case-insensitive comparisons,
' use an instance of CaseInsensitiveComparer class.</p>
' <p>In order for objects to be accepted in the comparison, they must implement the
' IComparable interface. The comparer will then call the object through the IComparable.Compare
' method, passing in the value for the object to compare itself with.<br><br>
' The return value should indicate the objects relation to the value passed in:<br>
' Positive number = object > value<br>
' Negative number = object < value<br>
' Return of zero  = object = value<br>
'
' @see ComparerStatic
'
Option Explicit
Implements IComparer

''
' Returns the comparison between to values.
'
' @param x The left-hand-side of the equation.
' @param y the right-hand-side of the equation.
' @return The relationship between the values.<br>
' Relationship guide:<br>
' x is equal to y returns 0.<br>
' x is greater than y returns a positive value (1).<br>
' x is less than y returns a negative value (-1).
'
Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long
    Compare = SZCompareVariants(x, y)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparer Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparer_Compare(x As Variant, y As Variant) As Long
    IComparer_Compare = SZCompareVariants(x, y)
End Function
