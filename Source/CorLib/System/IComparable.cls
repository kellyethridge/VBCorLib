VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IComparable"
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
' Module: IComparable
'

''
' Defines a generalized type-specific comparison method that a value type or class implements to order or sort its instances.
'
' @remarks This interface is implemented by types whose values can be ordered or sorted. It requires that implementing types define a single method,
' <a href="IComparable.CompareTo.html">CompareTo(Value)</a>, that indicates whether the position of the current instance in the sort order is
' before, after, or the same as a second object of the same type. The instance's <b>IComparable</b> implementation is called automatically by methods such
' as <a href="CorArray.Sort.html">CorArray.Sort</a> and <a href="ArrayList.Sort.html">ArrayList.Sort</a>.
' <p>The implementation of the <a href="IComparable.CompareTo.html">CompareTo(Value)</a> method must return a 32-bit integer that has one of three values, as shown in the following table.</p>
' <table style="width: auto">
'   <tr>
'       <th>Value</th><th>Meaning</th>
'   </tr>
'   <tr>
'       <td>Less than zero</td><td>The current instance precedes the value specified by the <a href="IComparable.CompareTo.html">CompareTo</a> method in the sort order.</td>
'   </tr>
'   <tr>
'       <td>Zero</td><td>This current instance occurs in the same position in the sort order as the value specified by the <a href="IComparable.CompareTo.html">CompareTo</a> method.</td>
'   </tr>
'   <tr>
'       <td>Greater than zero</td><td>This current instance follows the object specified by the <a href="IComparable.CompareTo.html">CompareTo</a> method in the sort order.</td>
'   </tr>
' </table>
' <p>All number types (such as <b>Long</b> and <b>Double</b>) are supported by the default <a href="Comparer.html">Comparer</a>, as are <b>String</b> and <b>Date</b>.
' Custom types should also provide their own implementation of <b>IComparable</b> to enable object instances to be ordered or sorted.</p>
' @include "..\DocIncludes\IComparable.html"
'@Folder("CorLib.System")
Option Explicit

''
' Compares the current instance with another object of the same type and returns an integer that indicates whether the current instance precedes, follows, or occurs in the same position in the sort order as the other object.
'
' @param Value A value to compare with this instance.
' @return A value that indicates the relative order of the objects being compared. The return value has these meanings:
' <table style="width: auto">
'   <tr>
'       <th>Value</th><th>Meaning</th>
'   </tr>
'   <tr>
'       <td>Less than zero</td><td>The current instance precedes the value specified by the <a href="IComparable.CompareTo.html">CompareTo</a> method in the sort order.</td>
'   </tr>
'   <tr>
'       <td>Zero</td><td>This current instance occurs in the same position in the sort order as the value specified by the <a href="IComparable.CompareTo.html">CompareTo</a> method.</td>
'   </tr>
'   <tr>
'       <td>Greater than zero</td><td>This current instance follows the object specified by the <a href="IComparable.CompareTo.html">CompareTo</a> method in the sort order.</td>
'   </tr>
' </table>
' @include "..\DocIncludes\IComparable.CompareTo.html"
Public Function CompareTo(ByRef Value As Variant) As Long: End Function

