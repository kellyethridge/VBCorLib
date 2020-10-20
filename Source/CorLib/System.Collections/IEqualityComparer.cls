VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEqualityComparer"
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
' Module: IEqualityComparer
'

''
' Defines methods to support the comparison of values for equality.
'
' @remarks <p>This interface allows the implementation of customized equality comparison for collections.
' That is, you can create your own definition of equality, and specify that this definition be used with a
' collection type that accepts the <b>IEqualityComparer</b> interface. In the VBCorLib Framework, constructor of the
' Hashtable collection type accepts this interface.</p>
' <p>This interface supports only equality comparisons. Customization of comparisons for sorting and ordering is provided by the <a href="IComparer.html">IComparer</a> interface.</p>
' <p>A default implementation can be accessed using <a href="EqualityComparerStatic.Get.Default.html">EqualityComparer.Default</a>.</p>
'@Folder("CorLib.System.Collections")
Option Explicit

''
' Determines whether the specified values are equal.
'
' @param x The first value to compare.
' @param y The second value to compare.
' @return <b>True</b> if the specified objects are equal; otherwise, <b>False</b>.
' @remarks <p>Implement this method to provide a customized equality comparison for values.</p>
' <p><b>Notes to Implementers</b></p>
' <p>The <b>Equals</b> method is reflexive, symmetric, and transitive. That is, it returns <b>True</b> if used to
' compare a value with itself; <b>True</b> for two values x and y if it is <b>True</b> for y and x; and <b>True</b> for
' two values x and z if it is <b>True</b> for x and y and also <b>True</b> for y and z.</p>
' <p>Implementations are required to ensure that if the <b>Equals</b> method returns true for two values x and y,
' then the value returned by the <a href="IEqualityComparer.GetHashCode.html">GetHashCode</a> method for x must equal the value returned for y.</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>x</i> and <i>y</i> are of different types and neither one can handle comparisons with the other.</td>
'   </tr>
' </table>
Public Function Equals(ByRef x As Variant, ByRef y As Variant) As Boolean: End Function

''
' Returns a hash code for the specified value.
'
' @param Obj The value for which a hash code is to be returned.
' @return A hash code for the specified object.
' @remarks <p>Implement this method to provide customized hash codes for values,corresponding to the customized equality comparison provided by the <a href="IEqualityComparer.Equals.html">Equals</a> method.</p>
' <p><b>Notes to Implementers</b></p>
' <p>Implementations are required to ensure that if the <a href="IEqualityComparer.Equals.html">Equals</a> method returns <b>True</b> for two values x and y, then the value returned by the <b>GetHashCode</b> method for x must equal the value returned for y.</p>
Public Function GetHashCode(ByRef Obj As Variant) As Long: End Function
