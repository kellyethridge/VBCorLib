VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderFallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
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
' Module: EncoderFallback
'

''
' Provides a failure-handling mechanism, called a fallback, for an input character that cannot be converted to an encoded output byte sequence.
'
' @see EncoerFallbackStatic
' @see EncoderExceptionFallback
' @see EncoderReplacementFallback
'@Folder("CorLib.System.Text")
Option Explicit

''
' When implemented in a derived class, gets the maximum number of characters the current <b>EncoderFallback</b> object can return.
'
' @return The maximum number of characters the current <b>EncoderFallback</b> object can return.
' @remarks For an exception fallback handler, this property should return zero.
'
Public Property Get MaxCharCount() As Long: End Property

''
' When implemented in a derived class, initializes a new instance of the EncoderFallbackBuffer class.
'
' @return An object that provides a fallback buffer for an encoder.
' @remarks This method returns an object derived from EncoderFallbackBuffer that is responsible for performing the fallback operation.
' For example, of the predefined <b>EncoderFallback</b> implementations in <b>VBCorLib</b>, the <b>EncoderReplacementFallback.CreateFallbackBuffer</b>
' method returns an <b>EncoderReplacementFallbackBuffer</b> object, and the <b>EncoderExceptionFallback.CreateFallbackBuffer</b> method returns an <b>EncoderExceptionFallbackBuffer</b> object.
'
Public Function CreateFallbackBuffer() As EncoderFallbackBuffer: End Function

''
' Determines whether the specified value passed in is equal to the current object.
'
' @param Value The value to compare to the current object.
' @return <b>True</b> if the specified value is equal to the current object; otherwise, <b>False</b>.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Serves as the default hash function.
'
' @return A hash code for the current object.
' @see IObject
'
Public Function GetHashCode() As Long: End Function

''
' Returns a string that represents the current object.
'
' @return A string the represents the current object.
' @see IObject
'
Public Function ToString() As String: End Function
