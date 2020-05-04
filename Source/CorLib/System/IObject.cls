VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IObject"
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
' Module: IObject
'

''
' This is the base interface to allow objects to be utilized throughout most of VBCorLib.
'
' @include "..\DocIncludes\IObject.html"
Option Explicit

''
' Returns a boolean indicating the equality of this instance versus the passed in value.
'
' @param value The value to compare this instance against for equality.
' @return Boolean indicating the equality of this instance and the value.
' @include "..\DocIncludes\IObject.Equals.html"
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number to identify this object instance.
'
' @return A 32-bit integer hashcode value.
' @remarks A hash code is a numeric value that is used to identify an object during equality testing. It can also serve as an index for an object in a collection.
' <p>The <b>GetHashCode</b> method is suitable for use in hashing algorithms and data structures such as a hash table.</p>
' <p>The default implementation of the <b>GetHashCode</b> method does not guarantee unique return values for different values. Furthermore, VBCorLib
' does not guarantee the default implementation of the <b>GetHashCode</b> method, and the value it returns will be the same between different versions of VBCorLib.
' Consequently, the default implementation of this method must not be used as a unique object identifier for hashing purposes.</p>
' <h5>Notes to Implementers</h5>
' A hash function is used to quickly generate a number (hash code) that corresponds to the value of an object. Hash functions are usually specific to each Type and, for uniqueness, must use at least one of the instance fields as input.
' A hash function must have the following properties:
' <ul>
'   <li>If two values compare as equal, the <b>GetHashCode</b> method for each value must return the same value. However, if two values
'       do not compare as equal, the <b>GetHashCode</b> methods for the two values do not have to return different values.</li>
'   <li>The <b>GetHashCode</b> method for an object must consistently return the same hash code as long as there is no modification to the object state that
'       determines the return value of the object's <a href="IObject.Equals.html">Equals</a> method. Note that this is true only for the current execution of
'       an application, and that a different hash code can be returned if the application is run again.</li>
'   <li>For the best performance, a hash function must generate a random distribution for all input.</li>
' </ul>
' <p>Implementations of the <b>GetHashCode</b> method must not result in circular references. For example, if <b>ClassA.GetHashCode</b> calls <b>ClassB.GetHashCode</b>, </b>ClassB.GetHashCode</b> must not call <b>ClassA.GetHashCode</b> either directly or indirectly.</p>
' <p>Implementations of the <b>GetHashCode</b> method must not throw exceptions.</p>
' @include "..\DocIncludes\IObject.GetHashCode.html"
Public Function GetHashCode() As Long: End Function

''
' Returns a string representation of the object.
'
' @return A string representation of the object.
' @remarks This should be overridden to return appropriate string values
' for the given object.As examples, the StringBuilder returns the current state of the
' internal string that is being built, and the CorDateTime object returns
' the date and time for that object instance.
' @include "..\DocIncludes\IObject.Equals.html"
Public Function ToString() As String: End Function

