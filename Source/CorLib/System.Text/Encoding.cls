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
' @return <b>True</b> if the current <a href="Encoding.html">Encoding</a> is read-only; otherwise, <b>False</b>. The default is <b>True</b>.
Public Property Get IsReadOnly() As Boolean: End Property

''
' Creates a shallow copy of the current <a href="Encoding.html">Encoding</a> object.
'
' @return A copy of the current <a href="Encoding.html">Encoding</a> object.
' @remarks The clone is writable even if the original <a href="Encoding.html">Encoding</a> object is read-only.
' Therefore, the properties of the clone can be modified.
' <p>A shallow copy of an object is a copy of the object only. If the object contains references to other objects,
' the shallow copy does not create copies of the referred objects. It refers to the original objects instead.
' In contrast, a deep copy of an object creates a copy of the object and a copy of everything directly or indirectly referenced by that object.</p>
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
' When implemented in a derived class, gets the human-readable description of the current encoding.
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
' @see #GetBytesEx
' @see #GetMaxByteCount
' @see #GetEncoder
' @include "..\DocIncludes\Encoding.GetByteCount.html"
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long: End Function

''
' Encodes a set of characters into an array of bytes.
'
' @param Chars The Integer array or String containing the set of characters to encode.
' @param Index The index of the first character to encode.
' @param Count The number of characters to encode.
' @return A byte array containing the results of encoding the specified set of characters.
' @see #GetBytesEx
' @see #GetEncoder
' @see #GetByteCount
' @see #GetMaxByteCount
' @include "..\DocIncludes\Encoding.GetBytes.html"
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
' @see #GetBytes
' @see #GetEncoder
' @see #GetByteCount
' @see #GetMaxByteCount
' @include "..\DocIncludes\Encoding.GetBytesEx.html"
Public Function GetBytesEx(ByRef Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long: End Function

''
' When implemented in a derived class, calculates the number of characters produced by decoding a sequence of bytes from the specified byte array.
'
' @param Bytes The byte array containing the sequence of bytes to decode.
' @param Index The index of the first byte to decode.
' @param Count The number of bytes to decode.
' @return The number of characters produced by decoding the specified sequence of bytes.
' @see #GetCharsEx
' @see #GetMaxCharCount
' @see #GetDecoder
' @include "..\DocIncludes\Encoding.GetCharCount.html"
Public Function GetCharCount(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long: End Function

''
' When implemented in a derived class, decodes all the bytes in the specified byte array into a set of characters.
'
' @param Bytes The byte array containing the sequence of bytes to decode.
' @param Index The index of the first byte to decode.
' @param Count The number of bytes to decode.
' @return A character array containing the results of decoding the specified sequence of bytes.
' @see #GetCharsEx
' @see #GetCharCount
' @see #GetMaxCharCount
' @see #GetDecoder
' @see #GetString
' @include "..\DocIncludes\Encoding.GetChars.html"
Public Function GetChars(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Integer(): End Function

''
' When implemented in a derived class, decodes a sequence of bytes from the specified byte array into the specified character array.
'
' @param Bytes The byte array containing the sequence of bytes to decode.
' @param ByteIndex The index of the first byte to decode.
' @param ByteCount The number of bytes to decode.
' @param Chars The character array to contain the resulting set of characters.
' @param CharIndex The index at which to start writing the resulting set of characters.
' @return The actual number of characters written into <i>Chars</i>.
' @see #GetChars
' @see #GetCharCount
' @see #GetMaxCharCount
' @see #GetDecoder
' @see #GetString
' @include "..\DocIncludes\Encoding.GetCharsEx.html"
Public Function GetCharsEx(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long: End Function

''
' When implemented in a derived class, obtains a decoder that converts an encoded sequence of bytes into a sequence of characters.
'
' @return A <a href="Decoder.html">Decoder</a> that converts an encoded sequence of bytes into a sequence of characters.
' @remarks The <a href="Decoder.GetChars.html">Decoder.GetChars</a> method converts sequential blocks of bytes into sequential
' blocks of characters, in a manner similar to the <a href="Encoding.GetChars.html">GetChars</a> method of this class.
' However, a <a href="Decoder.html">Decoder</a> maintains state information between calls so it correctly
' decodes byte sequences that span blocks. The <a href="Decoder.html">Decoder</a> also preserves trailing bytes at the end of
' data blocks and uses the trailing bytes in the next decoding operation. Therefore, <b>GetDecoder</b>
' and <a href="Encoding.GetEncoder.html">GetEncoder</a> are useful for network transmission and file operations, because those
' operations often deal with blocks of data instead of a complete data stream.
' @see Decoder
' @see #GetChars
' @see #GetString
' @see #GetCharCount
' @see #GetEncoder
Public Function GetDecoder() As Decoder: End Function

''
' When implemented in a derived class, obtains an encoder that converts a sequence of Unicode characters into an encoded sequence of bytes.
'
' @return A <a href="Encoder.html">Encoder</a> that converts a sequence of Unicode characters into an encoded sequence of bytes.
' @remarks The <a href="Encoder.GetBytes.html">Encoder.GetBytes</a> method converts sequential blocks of characters into sequential
' blocks of bytes, in a manner similar to the GetBytes method of this class. However, an <a href="Encoder.html">Encoder</a> maintains
' state information between calls so it correctly encodes character sequences that span blocks. The <a href="Encoder.html">Encoder</a>
' also preserves trailing characters at the end of data blocks and uses the trailing characters in the next encoding operation. For example,
' a data block might end with an unmatched high surrogate, and the matching low surrogate might be in the next data block.
' Therefore, <a href="Encoding.GetDecoder.html">GetDecoder</a> and <b>GetEncoder</b> are useful for network transmission and file operations,
' because those operations often deal with blocks of data instead of a complete data stream.
' @see Encoder
' @see #GetBytes
' @see #GetByteCount
' @see #GetDecoder
Public Function GetEncoder() As Encoder: End Function

''
' Gets the <a href="DecoderFallback.html">DecoderFallback</a> object for the current <a href="Encoding.html">Encoding</a> object.
'
' @return The decoder fallback object for the current <a href="Encoding.html">Encoding</a> object.
' @see DecoderFallback
' @see DecoderFallBackBuffer
' @include "..\DocIncludes\Encoding.Get.DecoderFallback.html"
Public Property Get DecoderFallback() As DecoderFallback: End Property

''
' Sets the <a href="DecoderFallback.html">DecoderFallback</a> object for the current <a href="Encoding.html">Encoding</a> object.
'
' @param Value The <a href="DecoderFallback.html">DecoderFallback</a> object to set.
' @see DecoderFallback
' @see DecoderFallBackBuffer
' @include "..\DocIncludes\Encoding.Set.DecoderFallback.html"
Public Property Set DecoderFallback(ByVal Value As DecoderFallback): End Property

''
' Gets the <a href="EncoderFallback.html">EncoderFallback</a> object for the current <a href="Encoding.html">Encoding</a> object.
'
' @Return The encoder fallback object for the current <a href="Encoding.html">Encoding</a> object.
' @see EncoderFallback
' @see EncoderFallbackBuffer
' @include "..\DocIncludes\Encoding.Get.EncoderFallback.html"
Public Property Get EncoderFallback() As EncoderFallback: End Property

''
' Sets the <a href="EncoderFallback.html">EncoderFallback</a> for the current <a href="Encoding.html">Encoding</a> object.
'
' @param Value The <a href="EncoderFallback.html">EncoderFallback</a> object to set.
' @see EncoderFallback
' @see EncoderFallbackBuffer
' @include "..\DocIncludes\Encoding.Set.EncoderFallback.html"
Public Property Set EncoderFallback(ByVal Value As EncoderFallback): End Property

''
' When implemented in a derived class, calculates the maximum number of bytes produced by encoding the specified number of characters.
'
' @param CharCount The number of characters to be encoded.
' @see #GetBytes
' @see #GetByteCount
' @include "..\DocIncludes\Encoding.GetMaxByteCount.html"
Public Function GetMaxByteCount(ByVal CharCount As Long) As Long: End Function

''
' Returns the maximum number of characters than can be decoded from the number of bytes specified.
'
' @param ByteCount The number of bytes to be decoded.
' @return The maximum number of characters that can be decoded from the specified number of bytes.
' @see #GetChars
' @see #GetString
' @see #GetCharCount
' @see #GetDecoder
' @include "..\DocIncludes\Encoding.GetMaxCharCount.html"
Public Function GetMaxCharCount(ByVal ByteCount As Long) As Long: End Function

''
' When implemented in a derived class, returns a sequence of bytes that specifies the encoding used.
'
' @return A byte array containg the bytes to be used to identify this encoding type.<br>
' -or-<br>
' A byte array of length zero, if a preamble is not required.
' @include "..\DocIncludes\Encoding.GetPreamble.html"
Public Function GetPreamble() As Byte(): End Function

''
' When implemented in a derived class, decodes all the bytes in the specified byte array into a string.
'
' @param Bytes The byte array containing the sequence of bytes to decode.
' @param Index The index of the first byte to decode.
' @param Count The number of bytes to decode.
' @return A String containing the results of decoding the specified sequence of bytes.
' @see #GetChars
' @see #GetDecoder
' @see #GetCharCount
' @see #GetMaxCharCount
' @include "..\DocIncludes\Encoding.GetString.html"
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
