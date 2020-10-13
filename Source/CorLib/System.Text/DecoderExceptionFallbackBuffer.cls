VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecoderExceptionFallbackBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: DecoderExceptionFallbackBuffer
'

''
' Throws <a href="DecoderFallbackException.html">DecoderFallbackException</a> when an encoded input byte sequence cannot be converted to a decoded output character.
'
Option Explicit
Implements IObject
Implements DecoderFallbackBuffer


''
' Throws <a href="DecoderFallbackException.html">DecoderFallbackException</a> when the input byte sequence cannot be decoded. The nominal return value is not used.
'
' @param BytesUnknown An input array of bytes.
' @param Index The index position of a byte in the input.
' @return None. No value is returned because the <b>Fallback</b> method always throws an exception.<br/>
' @remarks The <a href="Encoding.GetBytes.html">GetBytes</a> and <a href="EncodingStatic.Convert.html">Convert</a> methods call <b>Fallback</b> if they encounter an unknown byte in their input.
' In response, the <b>Fallback</b> method always throws <a href="DecoderFallbackException.html">DecoderFallbackException</a> and displays the input data.
' The <b>Fallback</b> method nominally indicates whether an exception is thrown if an input byte sequence cannot be decoded.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="DecoderFallbackException.html">DecoderFallbackException</a></td><td>This method always throws an exception that reports the value and index position of the input byte that cannot be decoded.</td>
'   </tr>
' </table>
Public Function Fallback(ByRef BytesUnknown() As Byte, ByVal Index As Long) As Boolean
    Dim Message As String
    
    Message = BuildMessage(BytesUnknown)
    
    Throw Cor.NewDecoderFallbackException2(Environment.GetResourceString(Argument_InvalidCodePageBytesIndex, Message, Index), BytesUnknown, Index)
End Function

''
' Retrieves the next character in the exception data buffer.
'
' @return The return value is always the Unicode character NULL (U+0000).
Public Function GetNextChar() As Integer
    GetNextChar = 0
End Function

''
' Causes the next call to <a href="DecoderExceptionFallbackBuffer.GetNextChar.html">GetNextChar</a> to access the exception data buffer character position that is prior to the current position.
'
' @return The return value is always <b>False</b>.
Public Function MovePrevious() As Boolean
    MovePrevious = False
End Function

''
' Gets the number of characters in the current <a href="DecoderExceptionFallbackBuffer.html">DecoderExceptionFallbackBuffer</a> object that remain to be processed.
'
' @return The return value is always zero.
Public Property Get Remaining() As Long
    Remaining = 0
End Property

''
' Determines whether the specified object is equal to the current object.
'
' @param Value The value to compare with the current object.
' @return <b>True</b> if the specified value is equal to the current object; otherwise, <b>False</b>.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Serves as the default hash function.
'
' @return A hash code for the current object.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string that represents the current object.
'
' @return A string that represents the current object.
Public Function ToString() As String
    ToString = MyBase.ToString(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildMessage(ByRef BytesUnknown() As Byte) As String
    Dim sb  As StringBuilder
    Dim i   As Long
    
    Set sb = StringBuilderCache.Acquire
    
    For i = LBound(BytesUnknown) To CorMath.Min(LBound(BytesUnknown), UBound(BytesUnknown))
        sb.AppendFormat "[{0:X2}]", BytesUnknown(i)
    Next
    
    If i <= UBound(BytesUnknown) Then
        sb.Append " ..."
    End If
    
    BuildMessage = StringBuilderCache.GetStringAndRelease(sb)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DecoderFallbackBuffer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DecoderFallbackBuffer_Equals(Value As Variant) As Boolean
    DecoderFallbackBuffer_Equals = Equals(Value)
End Function

Private Function DecoderFallbackBuffer_Fallback(BytesUnknown() As Byte, ByVal Index As Long) As Boolean
    DecoderFallbackBuffer_Fallback = Fallback(BytesUnknown, Index)
End Function

Private Function DecoderFallbackBuffer_GetHashCode() As Long
    DecoderFallbackBuffer_GetHashCode = GetHashCode
End Function

Private Function DecoderFallbackBuffer_GetNextChar() As Integer
    DecoderFallbackBuffer_GetNextChar = GetNextChar
End Function

Private Function DecoderFallbackBuffer_MovePrevious() As Boolean
    DecoderFallbackBuffer_MovePrevious = MovePrevious
End Function

Private Property Get DecoderFallbackBuffer_Remaining() As Long
    DecoderFallbackBuffer_Remaining = Remaining
End Property

Private Sub DecoderFallbackBuffer_Reset()
    Reset
End Sub

Private Function DecoderFallbackBuffer_ToString() As String
    DecoderFallbackBuffer_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
