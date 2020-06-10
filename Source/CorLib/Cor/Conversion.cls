VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conversion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2020 Kelly Ethridge
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
' Module: Conversion
'

''
' Provides methods to convert to specified data-types.
'
Option Explicit

''
' Converts a value to a CorDateTime object.
'
' @param Value The value to be converted to a<a href="CorDateTime.html">CorDateTime</a> object.
' @return The converted value.
' @remark This is a short-cut version of <a href="Convert.ToDateTime.html">Convert.ToDateTime</a>.
Public Function CDateTime(ByRef Value As Variant) As CorDateTime
    Set CDateTime = Convert.ToDateTime(Value)
End Function

''
' Converts a value to a 64-bit signed integer.
'
' @param Value The value to be converted
' @return Returns a 64-bit signed integer.
' @see Int64Static
' @see Constructors
' @include "..\DocIncludes\Conversion.CInt64.html"
Public Function CInt64(ByRef Value As Variant) As Int64
    CInt64 = ConversionMethods.CInt64(Value)
End Function

''
' Attempts to convert a value to a valid <a href="BigInteger.html">BigInteger</a> instance.
'
' @param value The value to be converted to <a href="BigInteger.html">BigInteger</a>.
' @return A new <a href="BigInteger.html">BigInteger</a> instance representing the converted value.
' @remarks The follow types are supported:<br/>
' <ul>
'   <li>String</li>
'   <li><a href="Int64Static.Int64.html">Int64</a></li>
'   <li>Long</li>
'   <li>Integer</li>
'   <li>Byte</li>
'   <li>Double</li>
'   <li>Single</li>
'   <li>Currency</li>
'   <li>Decimal</li>
'   <li>Byte Array</li>
'   <li><a href="BigInteger.html">BigInteger</a></li>
'   <li><b>Nothing</b></li>
' </ul>
' <pre>
' Dim b As BigInteger
' Set b = CBigInt(12345)
' </pre>
Public Function CBigInt(ByRef Value As Variant) As BigInteger
    Set CBigInt = ConversionMethods.CBigInt(Value)
End Function

''
' Returns the controlling IUnknown interface of the object.
'
' @param Obj The object get the controlling interface of.
' @return The controlling IUnknown interface.
Public Function CUnk(ByVal Obj As IUnknown) As IUnknown
    Set CUnk = Obj
End Function

''
' Returns the instance as type Object.
'
' @param Obj The instance to convert.
' @return Returns the instance as type Object.
Public Function CObj(ByVal Obj As Object) As Object
    Set CObj = Obj
End Function

