VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecoderReplacementFallbackBuffer"
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
' Module: DecoderReplacementFallbackBuffer
'

''
' Represents a substitute output string that is emitted when the original input byte sequence cannot be decoded. This class cannot be inherited.
'
'@Folder("CorLib.System.Text")
Option Explicit
Implements IObject
Implements DecoderFallbackBuffer

Private mFallbackChars()    As Integer
Private mFallbackLength     As Long
Private mFallbackCount      As Long
Private mFallbackIndex      As Long


''
' Prepares the replacement fallback buffer to use the current replacement string.
'
' @param BytesUnknown An input byte sequence. This parameter is ignored unless an exception is thrown.
' @param Index The index position of the byte in <i>BytesUnknown</i>. This parameter is ignored in this operation.
' @return <b>True</b> if the replacement string is not empty; <b>False</b> if the replacement string is empty.
' @remarks The <a href="Encoding.GetBytes.html">GetBytes</a> and <a href="Encoder.Convert.html">Convert</a> methods call <b>Fallback</b> if they encounter an unknown input byte sequence that cannot be decoded. If the return value of <b>Fallback</b> is <b>True</b>, the calling method can invoke the <a href="DecoderReplacementFallbackBuffer.GetNextChar.html">GetNextChar</a> method to obtain each character of the fallback buffer.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>This method is called again before the <a href="DecoderReplacementFallbackBuffer.GetNextChar.html">GetNextChar</a> method has read all the characters in the replacement fallback buffer.</td>
'   </tr>
' </table>
Public Function Fallback(ByRef BytesUnknown() As Byte, ByVal Index As Long) As Boolean
    If mFallbackCount > 0 Then _
        ThrowLastBytesRecursive BytesUnknown

    If mFallbackLength > 0 Then
        mFallbackCount = mFallbackLength
        mFallbackIndex = -1
        Fallback = True
    End If
End Function

Private Sub ThrowLastBytesRecursive(ByRef BytesUnknown() As Byte)
    Dim sb  As StringBuilder
    Dim i   As Long
    
    Set sb = StringBuilderCache.Acquire
    
    For i = LBound(BytesUnknown) To CorMath.Min(LBound(BytesUnknown) + 20, UBound(BytesUnknown))
        If sb.Length > 0 Then
            sb.AppendChar vbSpaceChar
        End If
        
        sb.AppendFormat "x{0:X2}", BytesUnknown(i)
    Next
    
    If i <= UBound(BytesUnknown) Then
        sb.Append " ..."
    End If
    
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_RecursiveFallbackBytes, StringBuilderCache.GetStringAndRelease(sb)), "BytesUnknown")
End Sub

''
' Retrieves the next character in the replacement fallback buffer.
'
' @return The next character in the replacement fallback buffer.
' @remarks The <a href="Encoding.GetBytes.html">Encoding.GetBytes</a> and <a href="Decoder.Convert.html">Decoder.Convert</a> methods call <a href="DecoderReplacementFallbackBuffer.Fallback.html">Fallback</a> if they encounter an unknown surrogate pair or character in their input. If the return value of <a href="DecoderReplacementFallbackBuffer.Fallback.html">Fallback</a> is <b>True</b>, the calling method can invoke <b>GetNextChar</b> to obtain each character of the replacement fallback buffer.
Public Function GetNextChar() As Integer
    mFallbackCount = mFallbackCount - 1
    mFallbackIndex = mFallbackIndex + 1
    
    ' the compiled version will have overflow checks off
    ' so we check here and let it continue.
    If mFallbackCount = &H7FFFFFFF Then
        mFallbackCount = -1
    End If
    
    If mFallbackCount >= 0 Then
        GetNextChar = mFallbackChars(mFallbackIndex)
    End If
End Function

''
' Causes the next call to <a href="DecoderReplacementFallbackBuffer.GetNextChar.html">GetNextChar</a> to access the character position in the replacement fallback buffer prior to the current character position.
'
' @return <b>True</b> if the <b>MovePrevious</b> operation was successful; otherwise, <b>False</b>.
Public Function MovePrevious() As Boolean
    If mFallbackCount >= -1 And mFallbackIndex >= 0 Then
        mFallbackCount = mFallbackCount + 1
        mFallbackIndex = mFallbackIndex - 1
        MovePrevious = True
    End If
End Function

''
' Gets the number of characters in the replacement fallback buffer that remain to be processed.
'
' @return The number of characters in the replacement fallback buffer that have not yet been processed.
Public Property Get Remaining() As Long
    If mFallbackCount > 0 Then
        Remaining = mFallbackCount
    End If
End Property

''
' Initializes all internal state information and data in the <a href="DecoderReplacementFallbackBuffer.html">DecoderReplacementFallbackBuffer</a> object.
'
' @remarks The <b>Reset</b> method discards any persisted state information and data that govern the emitting of a replacement string. Unpredictable results occur if the decoding operation continues after the <b>Reset</b> method is invoked.
Public Sub Reset()
    mFallbackCount = -1
    mFallbackIndex = -1
End Sub

''
' Determines whether the specified object is equal to the current object.
'
' @param Value The object to compare with the current object.
' @return <b>True</b> if the specified object is equal to the current object; otherwise, <b>False</b>.
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
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Fallback As DecoderReplacementFallback)
    mFallbackChars = CorString.ToCharArray(Fallback.DefaultString)
    mFallbackLength = Len(Fallback.DefaultString)
End Sub

Private Sub Class_Initialize()
    Reset
End Sub


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
