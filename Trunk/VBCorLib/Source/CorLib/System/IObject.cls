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
'    Module: IObject
'

''
' This is the base interface to allow objects to be utilized throughout most of VBCorLib.
'
' @remarks The <b>cObject</b> interface is utilized by many of the
' classes in VBCorLib. By implementing this interface in new classes, those
' classes help VBCorLib to perform specific tasks better. Such tasks include
' searching for a value.
' <p>If this interface is implemented, then the same methods should be
' implemented in the normal public interface to maintain consistency.</p>
' <p>If a class is not expected to be used by VBCorLib, then implementing
' this interface is not necessary.</p>
'
Option Explicit

''
' Returns a boolean indicating the equality of this instance versus the
' passed in value.
'
' @param value The value to compare this instance against for equality.
' @return Boolean indicating the equality of this instance and the value.
' @remarks <p>This should be overridden to provide the type of equality
' checking that is appropriate for each class.  Most classes simply
' check to see if the value is the same object instance as this object,
' meaning, they are both the same object in memory. cDateTime checks to
' see if the value has the same date as this instance, and TimeSpan
' checks to see if the value has the same number of milliseconds as this
' instance to determine equality.</p>
' <p>This method should not throw an exception if the wrong datatype is
' supplied. Instead, it should just return <i>False</i> to indicate that
' the value is not equal to this object instance.</p>
' <p><b>What should Equals when</b><br>
' <table>
'   <tr>
'       <th width="50%">What should Equals return when</td><th>Returns</td>
'   </tr>
'   <tr>
'       <td>Two Variables Point to the same object</td><td>True</td>
'   </tr>
'   <tr>
'       <td>Two objects represent the same internal value</td><td>True</td>
'   </tr>
' </table>
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number to identify this object instance.
'
' @return The number representing this instance.
' @remarks <p>A hashcode is used by VBCorLib to help distinguish different
' instances of objects. The Hashtable uses an object's hashcode to help
' search for a specific object. Since there is a finite set of numbers, each
' hashcode is not truly unique. A generated set of hashcodes should produce
' an evenly distributed set of values for a specific class type.</p>
' <p>This should be overridden to provide a value to identify
' this object instance. Most objects return their location in memory,
' unless the objects need to be considered the same as other objects.
' This occurs for both cDateTime and TimeSpan objects for example. All cDateTime
' objects that have the same date will generate the same hashcode. Same
' for TimeSpan. If the spans are the same, so are the hashcodes.</p>
' <p>A default implementation for this method could be<br>
' <pre>
' Public Function GetHashCode() As Long
'     GetHashCode = ObjPtr(CUnk(Me))
' End Function
' </pre>
'
'
Public Function GetHashCode() As Long: End Function

''
' Returns a string representation of the object.
'
' @return A string representation of the object.
' @remarks This should be overridden to return appropriate string values
' for the given object.As examples, the StringBuilder returns the current state of the
' internal string that is being built, and the cDateTime object returns
' the date and time for that object instance.
'
Public Function ToString() As String: End Function

