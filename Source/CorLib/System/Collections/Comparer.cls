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
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: DefaultComparer
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
    Compare = CompareVariants(x, y)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparer Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparer_Compare(x As Variant, y As Variant) As Long
    IComparer_Compare = CompareVariants(x, y)
End Function
