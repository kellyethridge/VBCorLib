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
' @remarks This class cannot be directly instantiated. To access the methods, use the variable name directly.
' <pre>
' Dim en As Encoding
' Set en = Encoding.UTF8
' </pre>
'
' @see Encoding
'
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


''
' Returns the default ANSI Encoding for the system.
'
' @return Default system encoding.
'
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
' @remarks This encoding is constructed using <a href="Constructors.NewUnicodeEncoding.html">Cor.NewUnicodeEncoding(True, True)</a>
' and may not be appropriate for your application.
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
' @remarks This encoding is constructed using <a href="Constructors.NewUnicodeEncoding.html">Cor.NewUnicodeEncoding(False, True)</a>
' and may not be appropriate for your application.
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
' @remarks This encoding is declared as New ASCIIEncoding.
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
' @remarks This encoding is declared as New UTF7Encoding.
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
' @remarks This encoding is constructed using <a href="Constructors.NewUTF8Encoding.html">Cor.NewUTF8Encoding(True)</a>
' and may not be appropriate for your application.
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
' @remarks The <a href="UTF32Encoding.html">UTF32Encoding</a> object that is returned by this property
' may not have the appropriate behavior for your app. It uses replacement fallback to replace each
' string that it cannot encode and each byte that it cannot decode with the Unicode REPLACEMENT CHARACTER (U+FFFE).
' Instead, you can call the <a href="Cor.NewUTF32Encoding.html">NewUTF32Encoding(Boolean, Boolean, Boolean)</a>
' constructor to instantiate a <a href="UTF32Encoding.html">UTF32Encoding</a> object whose fallback is either an
' <a href="EncoderFallbackException.html">EncoderFallbackException</a> or a <a href="DecoderFallbackException.html">DecoderFallbackException</a>.
Public Property Get UTF32() As UTF32Encoding
    If mUTF32 Is Nothing Then
        Set mUTF32 = Cor.NewUTF32Encoding(False, True)
    End If
    
    Set UTF32 = mUTF32
End Property

''
' Converts a set of bytes from one encoding to another encoding.
'
' @param srcEncoding The encoding to be used to decode the bytes.
' @param dstEncoding The encoding to be used to re-encode the decoded characters.
' @param Bytes The bytes to be converted from one encoding to the other.
' @param Index The starting byte to begin decoding at.
' @param Count The number of bytes to be included in the decoding process.
' @return The encoded bytes.
'
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
' Returns an encoding for a specified codepage.
'
' @param CodePageOrName The codepage number or name to encode with or 0 (zero) to use the default encoding.
' @param EncoderFallback An object that provides an error-handling procedure when a character cannot be encoded with the current encoding.
' @param DecoderFallback An object that provides an error-handling procedure when a byte sequence cannot be decoded with the current encoding.
' @return An encoding for a specific codepage or name.
Public Function GetEncoding(ByRef CodePageOrName As Variant, Optional ByVal EncoderFallback As EncoderFallback, Optional ByVal DecoderFallback As DecoderFallback) As Encoding
    Dim CodePage As Long
    
    Select Case VarType(CodePageOrName)
        Case vbString
            CodePage = mEncodingTable.GetCodePage(CStr(CodePageOrName))
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
            If mEncodings.Contains(CodePage) Then
                Set GetEncoding = mEncodings(CodePage)
            Else
                Set GetEncoding = NewCodePageEncoding(CodePage)
                mEncodings.Add CodePage, GetEncoding
            End If
    End Select
    
    If Not EncoderFallback Is Nothing Or Not DecoderFallback Is Nothing Then
        Set GetEncoding = GetEncoding.Clone
        Set GetEncoding.EncoderFallback = IfObject(EncoderFallback, GetEncoding.EncoderFallback)
        Set GetEncoding.DecoderFallback = IfObject(DecoderFallback, GetEncoding.DecoderFallback)
    End If
End Function

''
' Returns a list of minimal information about each encoding.
'
' @return An array of classes containing a minimal of information.
'
Public Function GetEncodings() As EncodingInfo()
    GetEncodings = mEncodingTable.GetEncodings
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

