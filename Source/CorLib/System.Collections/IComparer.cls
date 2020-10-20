VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IComparer"
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
' Module: IComparer
'

''
' Exposes a method that compares two values.
'
' @remarks This interface is used in conjunction with the <a href="CorArray.Sort.html">CorArray.Sort</a> and <a href="CorArray.BinarySearch.html">CorArray.BinarySearch</a>
' methods. It provides a way to customize the sort order of a collection. See the <a href="IComparer.Compare.html">Compare</a> method for notes on parameters and return value.
' <p>The default implementation of this interface is the <a href="Comparer.html">Comparer</a> class.</p>
' @see Comparer
' @include "..\DocIncludes\IComparer.html"
'@Folder("CorLib.System.Collections")
Option Explicit

''
' Compares two values and returns a value indicating whether one is less than, equal to, or greater than the other.
'
' @param a The first value to compare.
' @param b The second value to compare.
' @return A signed integer that indicates the relative values of <i>a</i> and <i>b</i>, as shown in the following table.
' <table style="width: auto">
'   <tr>
'       <th>Value</th><th>Meaning</th>
'   </tr>
'   <tr>
'       <td>Less than zero</td><td><i>a</i> is less than <i>b</i>.</td>
'   </tr>
'   <tr>
'       <td>Zero</td><td><i>a</i> equals <i>b</i>.</td>
'   </tr>
'   <tr>
'       <td>Greater than zero</td><td><i>a</i> is greater than <i>b</i>.</td>
'   </tr>
' </table>
' @remarks Comparing <b>Null</b> with any type is allowed and does not generate an exception when using <a href="IComparable.html">IComparable</a>. When sorting, <b>Null</b> is considered to be less than any other value.
' @see IComparable
' @include "..\DocIncludes\IComparer.Compare.html"
Public Function Compare(ByRef a As Variant, ByRef b As Variant) As Long: End Function

