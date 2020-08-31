VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Encoding"
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
' Module: Encoding
'

''
' Provides an interface for a provider to encode and decode unicode characters
' to and from bytes. Also contains information on the ability to use the
' encoded characters in certain situations without integerity loss.
'
' @see EncodingStatic
' @see ASCIIEncoding
' @see UTF7Encoding
' @see UTF8Encoding
' @see UTF32Encoding
' @see UnicodeEncoding
' @include "..\DocIncludes\Encoding.html"
Option Explicit


''
' When implemented in a derived class, gets a value indicating whether the current encoding is read-only.
'
' @return <b>true</b> if the current <a href="Encoding.html">Encoding</a> is read-only; otherwise, <b>false</b>. The default is <b>true</b>.
'
Public Property Get IsReadOnly() As Boolean: End Property

''
' Returns a shallow copy of the current <a href="Encoding.html">Encoding</a> object.
'
' @return A copy of the <a href="Encoding.html">Encoding</a> object.
' @remarks <p>An <a href="Encoding.html">Encoding</a> object is read-only by default.
' A clone of the <a href="Encoding.html">Encoding</a> object will become writable.</p>
'
Public Function Clone() As Encoding: End Function

''
' Returns the encoding name to be used in with the mail agent body tags.
'
' @return The body tag compatible encoding name.
'
Public Property Get BodyName() As String: End Property

''
' Returns the code page identifier for this encoding.
'
' @return Code page identifier.
'
Public Property Get CodePage() As Long: End Property

''
' When overridden in a derived class, gets the human-readable description of the current encoding.
'
' @include "..\DocIncludes\Encoding.EncodingName.html"
Public Property Get EncodingName() As String: End Property

''
' Returns the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
'
Public Property Get HeaderName() As String: End Property

''
' Indicates if this encoding can be used by browsers to display text.
'
' @return The indication for browser display capable.
'
Public Property Get IsBrowserDisplay() As Boolean: End Property

''
' Indicates if this encoding can be used to save data with this encoding.
'
' @return The indication for the browser saving capable.
'
Public Property Get IsBrowserSave() As Boolean: End Property

''
' Indicates if this encoding can be used to display mail and news by
' mail and news clients.
'
' @return Indication for mail and news client capable.
'
Public Property Get IsMailNewsDisplay() As Boolean: End Property

''
' Indicates if this encoding can be used to save date by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
'
Public Property Get IsMailNewsSave() As Boolean: End Property

''
' Returns the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
'
Public Property Get WebName() As String: End Property

''
' Returns the Windows Operating System's code page for this encoding.
'
' @return The Windows code page for this encoding.
'
Public Property Get WindowsCodePage() As Long: End Property

''
' Returns the number of bytes that would be produced from the set of characters using this encoding.
'
' @param Chars The Integer() or String containing the set of characters to encode.
' @param Index The index of the first character to encode.
' @param Count The number of characters to encode.
' @return The number of bytes produced by encoding the specified of characters.
' @include "..\DocIncludes\Encoding.GetByteCount.html"
' @see #GetBytesEx
' @see #GetMaxByteCount
' @see #GetEncoder
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long: End Function

''
' Encodes a set of characters into an array of bytes.
'
' @param Chars The Integer array or String containing the set of characters to encode.
' @param Index The index of the first character to encode.
' @param Count The number of characters to encode.
' @return A byte array containing the results of encoding the specified set of characters.
' @include "..\DocIncludes\Encoding.GetBytes.html"
' @see Encoding
' @see #GetBytesEx
' @see #GetEncoder
' @see #GetByteCount
' @see #GetMaxByteCount
Public Function GetBytes(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte(): End Function

''
' Encodes a set of characters into an array of bytes, returning the number of bytes produced.
'
' @param Chars The Integer array or String containing the set of characters to encode.
' @param CharIndex The index of the first character to encode.
' @param CharCount The number of characters to encode.
' @param Bytes The byte array to contain the resulting sequence of bytes.
' @param ByteIndex The index as which to start writing the resulting sequence of bytes.
' @return The number of bytes produce by the encoding.
' @include "..\DocIncludes\Encoding.GetBytesEx.html"
' @see Encoding
' @see #GetBytes
' @see #GetEncoder
' @see #GetByteCount
' @see #GetMaxByteCount
Public Function GetBytesEx(ByRef Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long: End Function

''
' When implemented in a derived class, calculates the number of characters produced by decoding a sequence of bytes from the specified byte array.
'
' @param Bytes The byte array containing the sequence of bytes to decode.
' @param Index The index of the first byte to decode.
' @param Count The number of bytes to decode.
' @return The number of characters produced by decoding the specified sequence of bytes.
' @include "..\DocIncludes\Encoding.GetCharCount.html"
' @see Encoding
' @see #GetCharsEx
' @see #GetMaxCharCount
' @see #GetDecoder
Public Function GetCharCount(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long: End Function

''
' Decodes a set of bytes into a set of characters.
'
' @param Bytes The set of bytes to be decoded into characters.
' @param Index The index of the first byte to be used in decoding.
' @param Count The number of bytes to be used in decoding.
' @return An array of characters decoded from the set of bytes.
'
Public Function GetChars(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Integer(): End Function

''
' Decodes a set of bytes into the supplied Integer array.
'
' @param Bytes The set of bytes to be decoded into characters.
' @param ByteIndex The index of the first byte to begin decoding from.
' @param ByteCount The number of bytes to be used in decoding.
' @param Chars The destination character array for the decoded bytes.
' @param CharIndex The first index to begin storing decoded bytes in <i>Chars</i>.
' @return The number of characters decoded from the array of bytes.
' @remarks The <i>Chars</i> array must be large enough to handle all the bytes that will
' be decoded. To ensure the <i>Chars</i> array is large enough, use either GetCharCount or
' GetMaxCharCount to determine a size that will hold the decoded bytes.
'
Public Function GetCharsEx(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long: End Function

''
' Returns a decoder that maintains state.
'
' @return The stateful decoder.
'
Public Function GetDecoder() As Decoder: End Function

''
' Returns an encoder that maintains state.
'
' @return The stateful encoder.
'
Public Function GetEncoder() As Encoder: End Function

''
' Returns the <a href="DecoderFallback.html">DecoderFallback</a> instance for the current Encoding object.
'
' @return An instance of a <a href="DecoderFallback.html">DecoderFallback</a>.
' @see Encoding
' @see DecoderFallBack
Public Property Get DecoderFallback() As DecoderFallback: End Property

''
' Sets the <a href="DecoderFallback.html">DecoderFallback</a> for the current Encoding object.
'
' @param Value The <a href="DecoderFallback.html">DecoderFallback</a> object to set.
' @remarks Assigning to a <i>read-only</i> Encoding object will cause an <a href="InvalidOperationException.html">InvalidOperationException</a> to be thrown.
'
Public Property Set DecoderFallback(ByVal Value As DecoderFallback): End Property

''
' Returns the <a href="EncodingFallback.html">EncodingFallback</a> for the current <a href="Encoding.html">Encoding</a> object.
'
' @Return The encoder fallback object for the current <a href="Encoding.html">Encoding</a> object.
' @see Encoding
' @see EncoderFallback
Public Property Get EncoderFallback() As EncoderFallback: End Property

''
' Sets the <a href="EncoderFallback.html">EncoderFallback</a> for the current Encoding object.
'
' @param Value The <a href="EncoderFallback.html">EncoderFallback</a> object to set.
' @remarks Assigning to a <i>read-only</i> Encoding object will cause an <a href="InvalidOperationException.html">InvalidOperationException</a> to be thrown.
'
Public Property Set EncoderFallback(ByVal Value As EncoderFallback): End Property

''
' Returns the maximum number of bytes that can be created from a specific number of characters.
'
' @param CharCount The number of characters to be encoded.
' @return The maximum number of bytes that can be generated from the specified number of characters.
'
Public Function GetMaxByteCount(ByVal CharCount As Long) As Long: End Function

''
' Returns the maximum number of characters than can be decoded from the number of bytes specified.
'
' @param ByteCount The number of bytes to be decoded.
' @return The maximum number of characters that can be decoded from the specified number of bytes.
'
Public Function GetMaxCharCount(ByVal ByteCount As Long) As Long: End Function

''
' Returns an array of bytes that represents this encoding.
'
' @return A byte array containg the bytes to be used to identify this encoding type.
' @remarks If no bytes are to be returned, then an empty Byte array should
' be returned, not a null array.
'
Public Function GetPreamble() As Byte(): End Function

''
' Decodes a set of bytes into a String.
'
' @param Bytes The set of bytes to be decoded into a string.
' @param Index The index of the first byte to be decoded.
' @param Count The number of bytes to be used in the decoding.
' @return A string containing the decoded set of bytes.
'
Public Function GetString(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As String: End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String: End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long: End Function

''
' Returns if the current encoding uses single-byte code points.
'
' @return Returns True for single-byte, False for multi-byte code points.
Public Property Get IsSingleByte() As Boolean: End Property
