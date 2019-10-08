VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Decoder"
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
' Module: Decoder
'

''
' Provides an interface to be used in converting blocks of bytes into unicode characters.
'
' @remarks Using a <b>Decoder</b> is different than using a normal <a href="Encoding.html">Encoding</a>
' because a <b>Decoder</b> can maintain state between calls. This allows a Decoder to decode characters
' that span method calls where as an <a href="Encoding.html">Encoding</a> method calls have
' no knowledge of previous data. A Decoder is useful in situation where unknown streams of data may need
' to be processed in multiple chunks such as over a network.
'
Option Explicit


''
' Returns a <a href="DecoderFallback.html">DecoderFallback</a> object for the current <a href="Decoder.html">Decoder</a> object.
'
' @return The current <a href="DecoderFallback.html">DecoderFallback</a> object.
Public Property Get Fallback() As DecoderFallback: End Property

''
' Sets a <a href="DecoderFallback.html">DecoderFallback</a> object for the current <a href="Decoder.html">Decoder</a> object.
'
' @param Value The <a href="DecoderFallback.html">DecoderFallback</a> object to set.
Public Property Set Fallback(ByVal Value As DecoderFallback): End Property

''
' Returns a <a href="DecoderFallbackBuffer.html">DecoderFallbackBuffer</a> object associated with the current <a href="Decoder.html">Decoder</a> object.
'
' @return The associated <a href="DecoderFallbackBuffer.html">DecoderFallbackBuffer</a> object.
Public Property Get FallbackBuffer() As DecoderFallbackBuffer: End Property

''
' Returns the number of characters a block of bytes can produce.
'
' @param Bytes The bytes to be used in decoding to characters.
' @param Index The starting index in <i>Bytes</i> to begin decoding from.
' @param Count The number of bytes to be used in the decoding process.
' @param Flush Tells the <a href="Decoder.html">Decoder</a> to clear any state after the call.
' @return The number of characters that can be produced from <i>Count</i> number of bytes.
' @remarks <p><b>GetCharCount</b> itself does not store state between method calls, but it will use
' any stored state from a previous call to <a href="Decoder.GetChars.html">GetChars</a>. Typically
' you would call this method immediately before calling <a href="Decoder.GetChars.html">GetChars</a>.</p>
' <p>Call with <i>Flush</i> equals <b>True</b> when then last chunk of data is being processed. This will
' ensure any characters that would normally have been stored as state for the next call will be processed
' during this call. No state is ever stored by this method, but any hanging characters need to be processed
' instead of being ignored.</p>
'
Public Function GetCharCount(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, Optional ByVal Flush As Boolean) As Long: End Function

''
' Returns the decoded unicode characters from a block of bytes.
'
' @param Bytes The bytes to be used in decoding to characters.
' @param ByteIndex The starting index to begin decoding bytes into characters.
' @param ByteCount The number of bytes to be used in the decoding process.
' @param Chars The output array for the decoded unicode characters.
' @param CharIndex The starting index in the array to begin storing characters.
' @param Flush Tells the <a href="Decoder.html">Decoder</a> to clear any state after the call.
' @return The number of characters actually decoded.
' @remarks <p><b>GetChars</b> converts an array of bytes into Unicode characters. Since Unicode
' characters could span a boundry between calls the method will store any state needed to resume
' converting characters in the next call. This allows for unknown chunks of data to be to
' converted without needed to enure each chunk contains complete characters. A typical scenario
' would be streaming data from a network.</p>
' <p>Because there is state stored between calls, the caller must inform the method that the last
' chunk of data is being processed so the method can correctly account for any state that might
' have persisted. This is achieved by passing <b>True</b> for the <i>Flush</i> parameter.
'
Public Function GetChars(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, Optional ByVal Flush As Boolean) As Long: End Function

''
' Sets <a href="Decoder.html">Decoder</a> to initial value.
'
' @remarks This methods clears any internal state of the <a href="Decoder.html">Decoder</a> object. The method clears any state information preserved from a previous call to GetChars or Convert, including trailing bytes at the end of the previous data block.
Public Sub Reset(): End Sub

''
' Converts an array of encoded bytes to UTF-16 encoded characters and stores the result in a character array.
'
' @param Bytes The byte array to convert.
' @param ByteIndex The starting position to convert in <i>Bytes</i>.
' @param ByteCount The number of bytes to be converted.
' @param Chars The char array to receive the converted characters.
' @param CharIndex The starting position to place converted characters.
' @param CharCount The maximum number of element of <i>chars</i> to use in the conversion.
' @param Flush <b>True</b> to indicate that no further data is to be converted; otherwise, <b>False</b>.
' @param BytesUsed When this method returns, contains the number of bytes that were used in the conversion. This parameter is passed uninitialized.
' @param CharsUsed When this method returns, contains the number of characters from chars that were produced by the conversion. This parameter is passed uninitialized.
' @param Completed When this method returns, contains <b>True</b> if all the characters specified by byteCount were converted; otherwise, <b>False</b>. This parameter is passed uninitialized.
' @remarks <p>The <b>Convert</b> method will attempt to convert up to <i>ByteCount</i> bytes but no more than would generate <i>CharCount</i> characters.</p>
' <p>An optimized implementation would attempt to convert as many bytes as are available to create <i>CharCount</i> characters. The default implementation
' simply tries to find a number of byte that when converted will fit in <i>CharCount</i> characters. If there are too many characters to fit then the number
' of bytes used to calculate the number of characters is halved. This does not lead to an optimal solution as <i>CharCount</i> will usually not be achieved.</p>
Public Sub Convert(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByVal Flush As Boolean, ByRef BytesUsed As Long, ByRef CharsUsed As Long, ByRef Completed As Boolean): End Sub

''
' Returns a string representation of this instance.
'
' @return A string that describes this instance.
'
Public Function ToString() As String: End Function

''
' Returns an indication of equality with the passed in value.
'
' @param value The value to test equality with.
' @param The result of the equality test.
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number that describes this instance.
'
' @return A pseudo-unique number that describes this instance.
'
Public Function GetHashCode() As Long: End Function

