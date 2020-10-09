VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2016 Kelly Ethridge
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
' Module: EncodingStatic
'

''
' Provides static methods used to retrieve existing encodings and convert between encodings.
'
' @see Encoding
' @see ASCIIEncoding
' @see UTF7Encoding
' @see UTF8Encoding
' @see UTF32Encoding
' @see UnicodeEncoding
' @include "..\DocIncludes\Encoding.html"
Option Explicit

Private mUTF32              As UTF32Encoding
Private mUTF8               As UTF8Encoding
Private mUTF7               As UTF7Encoding
Private mASCII              As ASCIIEncoding
Private mUnicode            As UnicodeEncoding
Private mBigEndianUnicode   As UnicodeEncoding
Private mEncodings          As New Hashtable
Private mEncodingTable      As New EncodingTable
Private mDefault            As Encoding
Private mEncodingInfos()    As EncodingInfo


''
' Gets an encoding for the operating system's current ANSI code page.
'
' @return An encoding for the operating system's current ANSI code page.
' @remarks <table style="width: auto">
'   <tr><th>Caution</th></tr>
'   <tr>
'       <td>Different computers can use different encodings as the default, and the default encoding can even change on a single computer.
'           Therefore, data streamed from one computer to another or even retrieved at different times on the same computer might be translated incorrectly.
'           In addition, the encoding returned by the <b>Default</b> property uses best-fit fallback to map unsupported characters to characters supported by the code page.
'           For these two reasons, using the default encoding is generally not recommended. To ensure that encoded bytes are decoded properly, your application
'           should use a Unicode encoding, such as <a href="UTF8Encoding.html">UTF8Encoding</a> or <a href="UnicodeEncoding.html">UnicodeEncoding</a>, with a preamble.
'           Another option is to use a higher-level protocol to ensure that the same format is used for encoding and decoding.</td>
'   </tr>
' </table>
' <p>The system ANSI code page defined by <b>Default</b> covers the ASCII set of characters, but the encoding is different from the encoding for ASCII.
' Because all <b>Default</b> encodings lose data, your application might use <a href="EncodingStatic.Get.UTF8.html">UTF8</a> instead. UTF-8 is often
' identical in the U+00 to U+7F range, but can encode other characters without loss.</p>
Public Property Get Default() As Encoding
    If mDefault Is Nothing Then
        Set mDefault = GetEncoding(0)
    End If
    
    Set Default = mDefault
End Property

''
' Gets an encoding for the UTF-16 format that uses the big endian byte order.
'
' @return An encoding object for the UTF-16 format that uses the big endian byte order.
' @include "..\DocIncludes\Encoding.Get.BigEndianUnicode.html"
Public Property Get BigEndianUnicode() As UnicodeEncoding
    If mBigEndianUnicode Is Nothing Then
        Set mBigEndianUnicode = Cor.NewUnicodeEncoding(True, True)
    End If
    
    Set BigEndianUnicode = mBigEndianUnicode
End Property

''
' Gets an encoding for the UTF-16 format using the little endian byte order.
'
' @return An encoding for the UTF-16 format using the little endian byte order.
' @include "..\DocIncludes\Encoding.Get.Unicode.html"
Public Property Get Unicode() As UnicodeEncoding
    If mUnicode Is Nothing Then
        Set mUnicode = Cor.NewUnicodeEncoding(False, True)
    End If
    
    Set Unicode = mUnicode
End Property

''
' Gets an encoding for the ASCII (7-bit) character set.
'
' @return An encoding for the ASCII (7-bit) character set.
' @include "..\DocIncludes\Encoding.Get.ASCII.html"
Public Property Get ASCII() As ASCIIEncoding
    If mASCII Is Nothing Then
        Set mASCII = New ASCIIEncoding
    End If
    
    Set ASCII = mASCII
End Property

''
' Gets an encoding for the UTF-7 format.
'
' @return An encoding for the UTF-7 format.
' @include "..\DocIncludes\Encoding.Get.UTF7.html"
Public Property Get UTF7() As UTF7Encoding
    If mUTF7 Is Nothing Then
        Set mUTF7 = New UTF7Encoding
    End If
    
    Set UTF7 = mUTF7
End Property

''
' Gets an encoding for the UTF-8 format.
'
' @return An encoding for the UTF-8 format.
' @include "..\DocIncludes\Encoding.Get.UTF8.html"
Public Property Get UTF8() As UTF8Encoding
    If mUTF8 Is Nothing Then
        Set mUTF8 = Cor.NewUTF8Encoding(True)
    End If
    
    Set UTF8 = mUTF8
End Property

''
' Gets an encoding for the UTF-32 format using the little endian byte order.
'
' @return An encoding object for the UTF-32 format using the little endian byte order.
' @include "..\DocIncludes\Encoding.Get.UTF32.html"
Public Property Get UTF32() As UTF32Encoding
    If mUTF32 Is Nothing Then
        Set mUTF32 = Cor.NewUTF32Encoding(False, True)
    End If
    
    Set UTF32 = mUTF32
End Property

''
' Converts a set of bytes from one encoding to another encoding.
'
' @param SrcEncoding The encoding of the source array, <i>Bytes</i>.
' @param DstEncoding The encoding of the output array.
' @param Bytes The array of bytes to convert.
' @param Index The index of the first element of <i>Bytes</i> to convert.
' @param Count The number of bytes to convert.
' @return An array of bytes containing the result of converting a range of bytes in bytes from <i>SrcEncoding</i> to <i>DstEncoding</i>.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td>
'        <td><i>SrcEncoding</i> is <b>Nothing</b>.<br>
'           -or-<br>
'           <i>DstEncoding</i> is <b>Nothing</b>.<br>
'           -or-<br>
'           <i>Bytes</i> is uninitialized.
'        </td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td><i>Index</i> and <i>Count</i> do not specify a valid range in the byte array.</td>
'   </tr>
'   <tr>
'       <td><a href="DecoderFallbackException.html">DecoderFallbackException</a></td>
'       <td>A fallback occurred<br>
'           -and-<br>
'           <b>SrcEncoding</b>.<a href="Encoding.Get.DecoderFallback.html">DecoderFallback</a> is set to <a href="DecoderExceptionFallback.html">DecoderExceptionFallback</a>.</td>
'   </tr>
'   <tr>
'       <td><a href="EncoderFallbackException.html">EncoderFallbackException</a></td>
'       <td>A fallback occurred<br>
'           -and-<br>
'           <b>DstEncoding</b>.<a href="Encoding.Get.EncoderFallback.html">EncoderFallback</a> is set to <a href="EncoderExceptionFallback.html">EncoderExceptionFallback</a>.</td>
'   </tr>
' </table>
Public Function Convert(ByVal SrcEncoding As Encoding, ByVal DstEncoding As Encoding, ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    If SrcEncoding Is Nothing Then _
        Error.ArgumentNull "SrcEncoding"
    If DstEncoding Is Nothing Then _
        Error.ArgumentNull "DstEncoding"
    
    Dim Ch() As Integer
    Ch = SrcEncoding.GetChars(Bytes, Index, Count)
    Convert = DstEncoding.GetBytes(Ch)
End Function

''
' Returns the encoding associated with the specified code page identifier or name. Optional parameters specify an error handler for characters that cannot be encoded and byte sequences that cannot be decoded.
'
' @param CodePageOrName The code page number or name to encode with or 0 (zero) to use the default encoding.
' @param EncoderFallback An object that provides an error-handling procedure when a character cannot be encoded with the current encoding.
' @param DecoderFallback An object that provides an error-handling procedure when a byte sequence cannot be decoded with the current encoding.
' @return An encoding for a specific codepage or name.
' @include "..\DocIncludes\Encoding.GetEncoding.html"
Public Function GetEncoding(ByRef CodePageOrName As Variant, Optional ByVal EncoderFallback As EncoderFallback, Optional ByVal DecoderFallback As DecoderFallback) As Encoding
    Dim CodePage As Long
    
    Select Case VarType(CodePageOrName)
        Case vbString
            CodePage = mEncodingTable.GetCodePage(CodePageOrName)
        Case vbLong, vbInteger, vbByte
            CodePage = CodePageOrName
        Case Else
            Throw Cor.NewArgumentException("A Code Page or Encoding Name is required.", "CodePageOrName")
    End Select
    
    If CodePage = 0 Then
        CodePage = GetACP
    End If
    
    Select Case CodePage
        Case 1200
            Set GetEncoding = Me.Unicode
        Case 1201
            Set GetEncoding = Me.BigEndianUnicode
        Case 20127
            Set GetEncoding = Me.ASCII
        Case 65000
            Set GetEncoding = Me.UTF7
        Case 65001
            Set GetEncoding = Me.UTF8
        Case 12000
            Set GetEncoding = Me.UTF32
        Case 12001
            Set GetEncoding = Cor.NewUTF32Encoding(True, True)
        Case Else
            Set GetEncoding = GetOrCreateEncoding(mEncodings(CodePage), CodePage)
    End Select
    
    If Not EncoderFallback Is Nothing Or Not DecoderFallback Is Nothing Then
        Set GetEncoding = GetEncoding.Clone
        Set GetEncoding.EncoderFallback = IfObject(EncoderFallback, GetEncoding.EncoderFallback)
        Set GetEncoding.DecoderFallback = IfObject(DecoderFallback, GetEncoding.DecoderFallback)
    End If
End Function

Private Function GetOrCreateEncoding(ByRef SearchResult As Variant, ByVal CodePage As Long) As Encoding
    If IsEmpty(SearchResult) Then
        Set GetOrCreateEncoding = NewCodePageEncoding(CodePage)
        mEncodings.Add CodePage, GetOrCreateEncoding
    Else
        Set GetOrCreateEncoding = SearchResult
    End If
End Function

''
' Returns a list of minimal information about each encoding.
'
' @return An array of classes containing a minimal of information.
'
Public Function GetEncodings() As EncodingInfo()
    If CorArray.IsNullOrEmpty(mEncodingInfos) Then
        mEncodingInfos = EncodingTable.GetEncodings
    End If
    
    GetEncodings = mEncodingInfos
End Function

''
' Returns reference to Encoding resources.
'
Friend Property Get EncodingTable() As EncodingTable
    Set EncodingTable = mEncodingTable
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NewCodePageEncoding(ByVal CodePage As Long) As Encoding
    Dim CodePageInfo As CPINFO
    
    If GetCPInfo(CodePage, CodePageInfo) = BOOL_FALSE Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingNotSupported, CodePage))
    
    Select Case CodePageInfo.MaxCharSize
        Case 1
            Set NewCodePageEncoding = Cor.NewSBCSCodePageEncoding(CodePage)
        Case 2
            Set NewCodePageEncoding = Cor.NewDBCSCodePageEncoding(CodePage)
    End Select
End Function

