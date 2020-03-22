VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringComparer"
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
' Module: StringComparer
'

''
' Represents string comparison operations based on a specific case or ordinal comparison rules.
'
' @remarks Implement this interface when a specific string comparison rule is necessary. The implementation
' should still allow the default processing of other value types other than string.
'
' @see StringComparerStatic
'
Option Explicit

''
' Compares two values relative to each other in an ordered set.
'
' @param x A value used in the comparison.
' @param y A value used in the comparison.
' @return Returns a numeric value indicating the relative order of the <i>x</i> parameter in relation to the <i>y</i> parameter.
' @remarks <p>If the <i>x</i> and <i>y</i> parameter values are considered equal then <b>zero</b> should be returned.
' If the <i>x</i> parameter is considered less than the <i>y</i> parameter in relative order then a
' <b>negative</b> value should be returned. Otherwise, the <i>x</i> parameter is considered greater
' than the <i>y</i> parameter and should return a positive value. There is no specific positive or
' negative value that is defined as the expected value, only the sign of the value is regarded.
' </p>
' <p>The example is the implementation used by the <a href="StringComparerStatic.Get.Ordinal.html">StringComparer.Ordinal</a> and <a href="StringComparerStatic.Get.OrdinalIgnoreCase.html">StringComparer.OrdinalIgnoreCase</a> implementation.</p>
' <pre>
'    Private Function StringComparer_Compare(x As Variant, y As Variant) As Long
'        If VarType(x) = vbString Then
'            If VarType(y) = vbString Then
'                StringComparer_Compare = CorString.Compare(x, y, mComparisonType)
'                Exit Function
'            End If
'        End If
'
'        StringComparer_Compare = Comparer.Default.Compare(x, y)
'    End Function
' </pre>
Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long: End Function

''
' Compares two values for equality.
'
' @param x A value to be compared for equality.
' @param y A value to be compared for equality.
' @return Returns <b>true</b> if <i>x</i> and <i>y</i> are equal; otherwise <b>false</b>.
' @remarks When implementing this method, any custom logic should be applied to only type <b>String</b>.
' All other value types should be passed to the default comparer for consistent behavior.
' <p>The example is the implementation used by the <a href="StringComparerStatic.Get.Ordinal.html">StringComparer.Ordinal</a> and <a href="StringComparerStatic.Get.OrdinalIgnoreCase.html">StringComparer.OrdinalIgnoreCase</a> implementation.</p>
' <pre>
'    Private mStringComparison As StringComparison
'
'    Private Function StringComparer_Equals(x As Variant, y As Variant) As Boolean
'        If VarType(x) = vbString Then
'            If VarType(y) = vbString Then
'                StringComparer_Equals = CorString.Equals(x, y, mStringComparison)
'                Exit Function
'            End If
'        End If
'
'        StringComparer_Equals = Object.Equals(x, y)
'    End Function
' </pre>
Public Function Equals(ByRef x As Variant, ByRef y As Variant) As Boolean: End Function

''
' Returns a partially unique numeric representation of the value.
'
' @param Obj The value to generate a hash code for.
' @return A partially unique number representing the value.
' @remarks Two values could generate the same hash code, but still not be equal. If two
' values generate the same hash code then additional comparisons should be done to check for equality.
' <p>The example is the implementation used by the <a href="StringComparerStatic.Get.Ordinal.html">StringComparer.Ordinal</a> and <a href="StringComparerStatic.Get.OrdinalIgnoreCase.html">StringComparer.OrdinalIgnoreCase</a> implementation.</p>
' <pre>
'    Private Function StringComparer_GetHashCode(Obj As Variant) As Long
'        If VarType(Obj) = vbString Then
'            If mStringComparison = OrdinalIgnoreCase Then
'                StringComparer_GetHashCode = Object.GetHashCode(UCase$(Obj))
'                Exit Function
'            End If
'        End If
'
'        StringComparer_GetHashCode = Object.GetHashCode(Obj)
'    End Function
' </pre>
Public Function GetHashCode(ByRef Obj As Variant) As Long: End Function
