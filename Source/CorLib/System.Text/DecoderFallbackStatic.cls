VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecoderFallbackStatic"
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
' Module: DecoderFallbackStatic
'
'@Folder("CorLib.System.Text")
Option Explicit

Private mReplacementFallback    As New DecoderReplacementFallback
Private mExceptionFallback      As New DecoderExceptionFallback


''
' Gets an object that outputs a substitute string in place of an input byte sequence that cannot be decoded.
'
' @returns A type derived from the DecoderFallback class.
' @remarks The default value is a DecoderReplacementFallback object that emits the QUESTION MARK character ("?", U+003F) in place of unknown byte sequences.
'
' @see DecoderReplacementFallback
' @see DecoderReplacementFallbackBuffer
'
Public Property Get ReplacementFallback() As DecoderFallback
    Set ReplacementFallback = mReplacementFallback
End Property


''
' Gets an object that throws an exception when an input byte sequence cannot be decoded.
'
' @returns A type derived from the DecoderFallback class.
' @remarks The default value is a DecoderExceptionFallback object.
'
' @see DecoderExceptionFallback
' @see DecoderExceptionFallbackBuffer
'
Public Property Get ExceptionFallback() As DecoderFallback
    Set ExceptionFallback = mExceptionFallback
End Property
