VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF8Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: UTF8Encoding
'
'            bytes   bits    UTF-8 representation
'            -----   ----    -----------------------------------
'            1        7      0vvvvvvv
'            2       11      110vvvvv 10vvvvvv
'            3       16      1110vvvv 10vvvvvv 10vvvvvv
'            4       21      11110vvv 10vvvvvv 10vvvvvv 10vvvvvv
'            -----   ----    -----------------------------------
'
'Surrogate:
'            Real Unicode value = (HighSurrogate - 0xD800) * 0x400 + (LowSurrogate - 0xDC00) + 0x10000
'

''
' A set of functions to be used to convert character arrays to and from byte arrays.
'
' @see Encoding
' @see ASCIIEncoding
' @see UTF7Encoding
' @see UTF32Encoding
' @see UnicodeEncoding
' @include "..\DocIncludes\UTF8Encoding.html"
Option Explicit
Implements IObject
Implements Encoding
Implements ICloneable

Private Const CODE_PAGE As Long = 65001

Private Const EmitUTF8IdentifierProp    As String = "EmitBOM"
Private Const IsReadOnlyProp            As String = "IsReadOnly"
Private Const DecoderFallbackProp       As String = "DecoderFallback"
Private Const EncoderFallbackProp       As String = "EncoderFallback"

Private Const FinalByte             As Long = &H40000000
Private Const Final2ndByte          As Long = FinalByte \ vbShift6Bits
Private Const Final3rdByte          As Long = FinalByte \ vbShift12Bits
Private Const Final4thByte          As Long = FinalByte \ vbShift18Bits
Private Const SupplimentarySeq      As Long = &H20000000
Private Const ThreeByteSeq          As Long = &H10000000
Private Const ThreeByteStartSeq     As Long = Final3rdByte Or ThreeByteSeq Or (ThreeByteSeq \ vbShift6Bits) Or (ThreeByteSeq \ vbShift12Bits)
Private Const FourByteStartSeq      As Long = Final4thByte Or SupplimentarySeq Or (SupplimentarySeq \ vbShift6Bits) Or (SupplimentarySeq \ vbShift12Bits) Or (SupplimentarySeq \ vbShift18Bits)

Private mIsReadOnly                 As Boolean
Private mShouldEmitUTF8Identifier   As Boolean
Private mEncoderFallback            As EncoderFallback
Private mDecoderFallback            As DecoderFallback
Private mFastLongsSA                As SafeArray1d
Private mFastLongsSAPtr             As Long
Private mFastLongs2SA               As SafeArray1d
Private mFastLongs2SAPtr            As Long
Private m1ByteUnknown(0)            As Byte
Private m2BytesUnknown(1)           As Byte
Private m3BytesUnknown(2)           As Byte


''
' When implemented in a derived class, gets a value indicating whether the current encoding is read-only.
'
' @return <b>True</b> if the current <a href="Encoding.html">Encoding</a> is read-only; otherwise, <b>False</b>. The default is <b>True</b>.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mIsReadOnly
End Property

''
' Creates a clone of the current Encoding instance.
'
' @return A cloned instance of a <a href="UTF8Encoding.html">UTF8Encoding</a> object.
' @remarks A cloned instance will have the read-only limitations removed.
'
Public Function Clone() As UTF8Encoding
    Set Clone = New UTF8Encoding
    Clone.InitClone Me
End Function

''
' Gets if the current encoding uses single-byte code points.
'
' @return Returns True for single-byte, False for multi-byte code points.
'
Public Property Get IsSingleByte() As Boolean
    IsSingleByte = False
End Property

''
' Gets the encoding name to be used in with the mail agent body tags.
'
' @return The body tag compatible encoding name.
'
Public Property Get BodyName() As String
    BodyName = Table.BodyName(CodePage)
End Property

''
' Gets the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
'
Public Property Get WebName() As String
    WebName = Table.WebName(CodePage)
End Property

''
' Gets the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
'
Public Property Get HeaderName() As String
    HeaderName = Table.HeaderName(CodePage)
End Property

''
' Gets the code page identifier for this encoding.
'
' @return Code page identifier.
'
Public Property Get CodePage() As Long
    CodePage = CODE_PAGE
End Property

''
' Gets the Windows Operating System's code page for this encoding.
'
' @return The Windows code page for this encoding.
'
Public Property Get WindowsCodePage() As Long
    WindowsCodePage = Table.WindowsCodePage(CodePage)
End Property

''
' Gets the human-readable description of the current encoding.
'
' @include "..\DocIncludes\Encoding.EncodingName.html"
Public Property Get EncodingName() As String
    EncodingName = Table.EncodingName(CodePage)
End Property

''
' Gets if this encoding can be used by browsers to display text.
'
' @return The indication for browser display capable.
'
Public Property Get IsBrowserDisplay() As Boolean
    IsBrowserDisplay = Table.IsBrowserDisplay(CodePage)
End Property

''
' Gets if this encoding can be used to save data with this encoding.
'
' @return The indication for the browser saving capable.
'
Public Property Get IsBrowserSave() As Boolean
    IsBrowserSave = Table.IsBrowserSave(CodePage)
End Property

''
' Gets if this encoding can be used to display mail and news by
' mail and news clients.
'
' @return Indication for mail and news client capable.
'
Public Property Get IsMailNewsDisplay() As Boolean
    IsMailNewsDisplay = Table.IsMailNewsDisplay(CodePage)
End Property

''
' Gets if this encoding can be used to save data by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
'
Public Property Get IsMailNewsSave() As Boolean
    IsMailNewsSave = Table.IsMailNewsSave(CodePage)
End Property

''
' Gets the <a href="DecoderFallback.html">DecoderFallback</a> object for the current <a href="Encoding.html">Encoding</a> object.
'
' @return The decoder fallback object for the current <a href="Encoding.html">Encoding</a> object.
' @see DecoderFallback
' @see DecoderFallbackBuffer
' @include "..\DocIncludes\Encoding.Get.DecoderFallback.html"
Public Property Get DecoderFallback() As DecoderFallback
    Set DecoderFallback = mDecoderFallback
End Property

''
' Sets the <a href="DecoderFallback.html">DecoderFallback</a> object for the current <a href="Encoding.html">Encoding</a> object.
'
' @param Value The decoder fallback object for the current <a href="Encoding.html">Encoding</a> object.
' @see DecoderFallback
' @see DecoderFallbackBuffer
' @include "..\DocIncludes\Encoding.Set.DecoderFallback.html"
Public Property Set DecoderFallback(ByVal Value As DecoderFallback)
    If mIsReadOnly Then _
        Error.InvalidOperation InvalidOperation_ReadOnly
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    
    Set mDecoderFallback = Value
End Property

''
' Gets the <a href="EncoderFallback.html">EncoderFallback</a> object for the current <a href="Encoding.html">Encoding</a> object.
'
' @return The encoder fallback object for the current <a href="Encoding.html">Encoding</a> object.
' @see EncoderFallback
' @see EncoderFallbackBuffer
' @include "..\DocIncludes\Encoding.Get.EncoderFallback.html"
Public Property Get EncoderFallback() As EncoderFallback
    Set EncoderFallback = mEncoderFallback
End Property

''
' Sets the <a href="EncoderFallback.html">EncoderFallback</a> object for the current <a href="Encoding.html">Encoding</a> object.
'
' @param Value The encoder fallback object for the current <a href="Encoding.html">Encoding</a> object.
' @see EncoderFallback
' @see EncoderFallbackBuffer
' @include "..\DocIncludes\Encoding.Set.EncoderFallback.html"
Public Property Set EncoderFallback(ByVal Value As EncoderFallback)
    If mIsReadOnly Then _
        Error.InvalidOperation InvalidOperation_ReadOnly
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    
    Set mEncoderFallback = Value
End Property

''
' Calculates the number of bytes produced by encoding the characters in the specified String or Integer().
'
' @param Chars The Integer array or String containing the set of characters to encode.
' @param Index The index of the first character to encode.
' @param Count The number of characters to encode.
' @return The number of bytes produced by encoding the specified set of characters.
' @include "..\DocIncludes\UTF8Encoding.GetByteCount.html"
' @see #GetBytesEx
' @see #GetMaxByteCount
' @see #GetEncoder
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim NullBytes()     As Byte
    Dim LocalChars()    As Integer
    Dim Range           As ListRange
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    Range = ValidateGetBytes(LocalChars, Index, Count)
    
    If Range.Count > 0 Then
        GetByteCount = GetByteCountCore(LocalChars, Range.Index, Range.Count, Nothing)
    End If
    
    FreeChars LocalChars
    Exit Function
    
Catch:
    FreeChars LocalChars
    ThrowOrErr Err
End Function

''
' Encodes all the characters in the specified character array or string into a sequence of bytes.
'
' @param Chars The Integer array or String containing the characters to encode.
' @param Index The index of the first character to begin encoding from.
' @param Count The number of characters to be encoded.
' @return A byte array containing the result of encoding the specified set of characters.
' @see #GetBytesEx
' @see #GetEncoder
' @see #GetByteCount
' @see #GetMaxByteCount
' @include "..\DocIncludes\UTF8Encoding.GetBytes.html"
Public Function GetBytes(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    Dim Range           As ListRange
    Dim LocalChars()    As Integer
    Dim ByteCount       As Long
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    Range = ValidateGetBytes(LocalChars, Index, Count)
    
    ByteCount = GetByteCountCore(LocalChars, Range.Index, Range.Count, Nothing)
    SAPtr(GetBytes) = SafeArrayCreateVectorEx(vbByte, 0, ByteCount)
    GetBytesCore LocalChars, Range.Index, Range.Count, GetBytes, 0, Nothing
    
    FreeChars LocalChars
    Exit Function
    
Catch:
    FreeChars LocalChars
    ThrowOrErr Err
End Function

''
' Encodes a set of characters into an array of bytes, returning the number of bytes produced.
'
' @param Chars The Integer array or String containing the characters to encode.
' @param CharIndex The index of the first character to begin encoding from.
' @param CharCount The number of characters to be encoded.
' @param Bytes The byte array to contain the resulting sequence of bytes.
' @param ByteIndex The index as which to start writing the resulting sequence of bytes.
' @return The number of bytes produce by the encoding.
' @see #GetBytes
' @see #GetEncoder
' @see #GetByteCount
' @see #GetMaxByteCount
' @include "..\DocIncludes\UTF8Encoding.GetBytesEx.html"
Public Function GetBytesEx(ByRef Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Dim LocalChars()    As Integer
    
    On Error GoTo Catch
    
    LocalChars = AsChars(Chars)
    ValidateGetBytesEx LocalChars, CharIndex, CharCount, Bytes, ByteIndex
    GetBytesEx = GetBytesCore(LocalChars, CharIndex, CharCount, Bytes, ByteIndex, Nothing)
    
    FreeChars LocalChars
    Exit Function
    
Catch:
    FreeChars LocalChars
    ThrowOrErr Err
End Function

''
' Calculates the number of characters produced by decoding a sequence of bytes from the specified byte array.
'
' @param Bytes The byte array containing the sequence of bytes to decode.
' @param Index The index of the first byte to decode.
' @param Count The number of bytes to decode.
' @return The number of characters produced by decoding the specified sequence of bytes.
' @see #GetCharsEx
' @see #GetMaxCharCount
' @see #GetDecoder
' @include "..\DocIncludes\UTF8Encoding.GetCharCount.html"
Public Function GetCharCount(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim Range As ListRange
    
    Range = ValidateGetChars(Bytes, Index, Count)
    
    GetCharCount = GetCharCountCore(Bytes, Range.Index, Range.Count, Nothing)
End Function

''
' Decodes a sequence of bytes from the specified byte array into a set of characters.
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
Public Function GetChars(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Integer()
    Dim Range       As ListRange
    Dim CharCount   As Long
    
    Range = ValidateGetChars(Bytes, Index, Count)
        
    CharCount = GetCharCountCore(Bytes, Range.Index, Range.Count, Nothing)
    SAPtr(GetChars) = SafeArrayCreateVectorEx(vbInteger, 0, CharCount)
    GetCharsCore Bytes, Range.Index, Range.Count, GetChars, 0, Nothing
End Function

''
' Decodes a sequence of bytes from the specified byte array into the specified character array.
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
' @include "..\DocIncludes\UTF8Encoding.GetCharsEx.html"
Public Function GetCharsEx(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    ValidateGetCharsEx Bytes, ByteIndex, ByteCount, Chars, CharIndex
    GetCharsEx = GetCharsCore(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Nothing)
End Function

''
' Decodes a range of bytes from a byte array into a string.
'
' @param Bytes The byte array containing the sequence of bytes to decode.
' @param Index The index of the first byte to decode.
' @param Count The number of bytes to decode.
' @return A String containing the results of decoding the specified sequence of bytes.
' @see #GetChars
' @see #GetCharsEx
' @see #GetDecoder
' @see #GetCharCount
' @see #GetMaxCharCount
' @include "..\DocIncludes\UTF8Encoding.GetString.html"
Public Function GetString(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As String
    Dim Range       As ListRange
    Dim CharCount   As Long
    Dim Chars()     As Integer
    
    Range = ValidateGetChars(Bytes, Index, Count)
    
    CharCount = GetCharCountCore(Bytes, Range.Index, Range.Count, Nothing)
    GetString = SysAllocStringLen(0, CharCount)
    Chars = AllocChars(GetString)
    
    On Error GoTo Catch
    GetCharsCore Bytes, Range.Index, Range.Count, Chars, 0, Nothing
    
    FreeChars Chars
    Exit Function
    
Catch:
    FreeChars Chars
    ThrowOrErr Err
End Function

''
' Returns a Unicode byte order mark encoded in UTF-8 format, if the constructor for this instance requests a byte order mark.
'
' @return A byte array containing the Unicode byte order mark, if the constructor for this instance requests a byte order mark. Otherwise, this method returns a byte array of length zero.
' @include "..\DocIncludes\UTF8Encoding.GetPreamble.html"
Public Function GetPreamble() As Byte()
    If mShouldEmitUTF8Identifier Then
        GetPreamble = Cor.NewBytes(&HEF, &HBB, &HBF)
    Else
        GetPreamble = Cor.NewBytes()     ' creates a 0-length array instead of a null array.
    End If
End Function

''
' Calculates the maximum number of bytes produced by encoding the specified number of characters.
'
' @param CharCount The number of characters to encode.
' @return The maximum number of bytes produced by encoding the specified number of characters.
' @see #GetBytesEx
' @see #GetByteCount
' @see #GetEncoder
' @include "..\DocIncludes\UTF8Encoding.GetMaxByteCount.html"
Public Function GetMaxByteCount(ByVal CharCount As Long) As Long
    If CharCount < 0 Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_NeedNonNegNum
    
    ' chars are a max 3 bytes each.
    GetMaxByteCount = (CharCount + 1) * 3 ' plus one incase the first char part of a surrogate pair from a previous operation.
End Function

''
' Calculates the maximum number of characters produced by decoding the specified number of bytes.
'
' @param ByteCount The number of bytes to decode.
' @return The maximum number of characters produced by decoding the specified number of bytes.
' @see #GetCharsEx
' @see #GetString
' @see #GetCharCount
' @see #GetDecoder
' @include "..\DocIncludes\UTF8Encoding.GetMaxCharCount.html"
Public Function GetMaxCharCount(ByVal ByteCount As Long) As Long
    If ByteCount < 0 Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_NeedNonNegNum
        
    GetMaxCharCount = ByteCount + 1 ' plus one incase the first byte is the last byte of a surrogate pair from a previous operation.
End Function

''
' Obtains an encoder that converts a sequence of Unicode characters into a UTF-8 encoded sequence of bytes.
'
' @return An <a href="Encoder.html">Encoder</a> that converts a sequence of Unicode characters into a UTF-8 encoded sequence of bytes.
' @see Encoder
' @see #GetBytesEx
' @see #GetByteCount
' @see #GetDecoder
' @include "..\DocIncludes\UTF8Encoding.GetEncoder.html"
Public Function GetEncoder() As Encoder
    Dim Encoder As New UTF8Encoder
    Encoder.Init Me
    Set GetEncoder = Encoder
End Function

''
' Obtains a decoder that converts a UTF-8 encoded sequence of bytes into a sequence of Unicode characters.
'
' @return A <a href="Decoder.html">Decoder</a> that converts a UTF-8 encoded sequence of bytes into a sequence of Unicode characters.
' @see Decoder
' @see #GetChars
' @see #GetCharsEx
' @see #GetString
' @see #GetCharCount
' @see #GetEncoder
' @include "..\DocIncludes\UTF8Encoding.GetDecoder.html"
Public Function GetDecoder() As Decoder
    Dim Decoder As New UTF8Decoder
    Decoder.Init Me
    Set GetDecoder = Decoder
End Function

''
' Returns a string representation of the current object.
'
' @include "..\DocIncludes\IObject.ToString.html"
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Determines whether the specified value is equal to the current <a href="UTF8Encoding.htm">UTF8Encoding</a> object.
'
' @param Value The value to compare with the current instance.
' @return <b>True</b> if value is an instance of <a href="UTF8Encoding.html">UTF8Encoding</a> and is equal to the current object; otherwise, <b>False</b>.
' @include "..\DocIncludes\UTF8Encoding.Equals.html"
Public Function Equals(ByRef Value As Variant) As Boolean
    Dim Other As UTF8Encoding
    
    On Error GoTo Catch
    Set Other = Value
    Equals = Other.EmitUTF8Identifier = mShouldEmitUTF8Identifier And _
        Other.DecoderFallback.Equals(mDecoderFallback) And _
        Other.EncoderFallback.Equals(mEncoderFallback)
        
Catch:
End Function

''
' Returns the hash code for the current instance.
'
' @return The hash code for the current instance.
' @include "..\DocIncludes\UTF8Encoding.GetHashCode.html"
Public Function GetHashCode() As Long
    GetHashCode = CODE_PAGE Xor _
                CLng(IIf(mShouldEmitUTF8Identifier, &H80000000, 0)) Xor _
                mEncoderFallback.GetHashCode Xor _
                mDecoderFallback.GetHashCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal ShouldEmitUTF8Identifier As Long, ByVal ThrowOnInvalidBytes As Boolean)
    mShouldEmitUTF8Identifier = ShouldEmitUTF8Identifier
    
    If ThrowOnInvalidBytes Then
        Set mEncoderFallback = Statics.EncoderFallback.ExceptionFallback
        Set mDecoderFallback = Statics.DecoderFallback.ExceptionFallback
    End If
End Sub

Friend Sub InitClone(ByVal Other As UTF8Encoding)
    Set mEncoderFallback = Other.EncoderFallback
    mShouldEmitUTF8Identifier = Other.EmitUTF8Identifier
    mIsReadOnly = False
End Sub

Private Sub Class_InitProperties()
    mFastLongsSA.cbElements = 4
    mFastLongsSA.cDims = 1
    mFastLongsSA.cElements = &H7FFFFFFF
    mFastLongsSAPtr = VarPtr(mFastLongsSA)
    
    mFastLongs2SA.cbElements = 4
    mFastLongs2SA.cDims = 1
    mFastLongs2SA.cElements = &H7FFFFFFF
    mFastLongs2SAPtr = VarPtr(mFastLongs2SA)
    
    Set mEncoderFallback = Cor.NewEncoderReplacementFallback(ChrW$(&HFFFD))
    Set mDecoderFallback = Cor.NewDecoderReplacementFallback(ChrW$(&HFFFD))
    mIsReadOnly = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friends
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get EmitUTF8Identifier() As Boolean
    EmitUTF8Identifier = mShouldEmitUTF8Identifier
End Property

Friend Function GetCharCountCore(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, ByVal Decoder As UTF8Decoder) As Long
    Dim MaxIndex        As Long
    Dim b               As Long
    Dim Char            As Long
    Dim Fallback        As DecoderFallbackBuffer
    Dim CharNoFlags     As Long
    Dim CharCount       As Long
    Dim CanFastLoop     As Boolean
    Dim BytesAddress    As Long
    Dim BytesLB         As Long
    Dim FastBytes()     As Long
    
    MaxIndex = Index + Count - 1
    
    If Not Decoder Is Nothing Then
        Char = Decoder.Char
    End If
    
    If Len1D(Bytes) > 0 Then
        BytesLB = LBound(Bytes)
        BytesAddress = VarPtr(Bytes(BytesLB))
        CanFastLoop = True
    End If
    
    Do While Index <= MaxIndex
        b = Bytes(Index)
        Index = Index + 1

        If Char = 0 Then
            ' Check if we have an ASCII character we can just stuff immediately.
            If b < &H80 Then
                CharCount = CharCount + 1
                
                ' Fast loop: adapted from .NET reference source.
                Dim AvailableBytes  As Long
                Dim MaxFastIndex    As Long
                Dim FastIndex       As Long
                
                If Not CanFastLoop Then
                    GoTo Continue
                End If
                
                AvailableBytes = MaxIndex - Index + 1
                
                ' Don't fall into the fast decoding loop if we don't have enough characters.
                ' Note that if we don't have enough bytes, MaxFastIndex will prevent us from entering the fast loop.
                If AvailableBytes <= 13 Then
                    ' Run through as many ASCII characters as quickly as we can here.
                    Do While Index <= MaxIndex
                        b = Bytes(Index)
                        Index = Index + 1
                        
                        ' Check if not an ASCII character.
                        If b > &H7F Then
                            GoTo ProcessChar
                        End If
                        
                        CharCount = CharCount + 1
                    Loop
                    
                    ' We got here without hitting a non-ascii character so
                    ' we are out of characters. Time to exit the loop.
                    Exit Do
                End If
                
                ' To compute the upper bound, assume that all characters are ASCII characters at this point,
                ' the boundary will be decreased for every non-ASCII character we encounter
                ' Also, we need 7 chars reserve for the unrolled ASCII decoding loop and for decoding of multibyte sequences
                MaxFastIndex = Index + AvailableBytes - 7
                
                Do While Index < MaxFastIndex
                    b = Bytes(Index)
                    Index = Index + 1
                    
                    ' Is it a non-ASCII character?
                    If b > &H7F Then
                        GoTo LongCode
                    End If
                    
                    CharCount = CharCount + 1
                    
                    ' Set up our faster access to the Bytes and Chars arrays.
                    mFastLongsSA.PVData = BytesAddress + (Index - BytesLB)
                    SAPtr(FastBytes) = mFastLongsSAPtr
                    FastIndex = 0
                    
                    ' We should be 4-byte aligned here so let'er rip!
                    ' We are working with 8 bytes at a time. Normally this
                    ' would require 8 byte reads and 8 char writes, but we
                    ' are accessing both arrays as Longs, so we get the
                    ' array access count down to 2 reads and 4 writes.
                    Do While Index < MaxFastIndex
                        Dim b2 As Long
                        b = FastBytes(FastIndex)
                        b2 = FastBytes(FastIndex + 1)
                        FastIndex = FastIndex + 2
                        
                        ' Check if any of the 8 bytes is a non-ASCII and bail if any are.
                        If (b Or b2) And &H80808080 Then
                            SAPtr(FastBytes) = vbNullPtr
                            GoTo LongCodeWithMask
                        End If
                        
                        CharCount = CharCount + 8
                        Index = Index + 8
                    Loop
                    
                    SAPtr(FastBytes) = vbNullPtr
                    Exit Do

LongCodeWithMask:
                    ' The multi-byte checks in previous code revealed that at least
                    ' one of the bytes is not ASCII, so here we are. Will do a quick
                    ' check to see if the first byte is at least ASCII and if so we
                    ' can store it and continue with the loop. This way there is a
                    ' chance that the multi-byte access didn't go to complete waste.
                    b = b And &HFF&
                    Index = Index + 1
                                        
                    If b <= &H7F Then
                        CharCount = CharCount + 1
                        GoTo ContinueFastLoop
                    End If
                
LongCode:
                    Dim CharC As Long
                    CharC = Bytes(Index)
                    Index = Index + 1
                    
                    If (b And &H40) = 0 Then
                        GoTo BadLongCode
                    ElseIf (CharC And &HC0) <> &H80 Then
                        GoTo BadLongCode
                    End If
                    
                    CharC = CharC And &H3F
                    
                    If b And &H20 Then
                        CharC = CharC Or ((b And &HF) * vbShift6Bits)
                        
                        If b And &H10 Then
                            ' 4-byte encoding (surrogate)
                            b = Bytes(Index)
                            
                            Dim SurrogateCheck As Long
                            SurrogateCheck = CharC \ vbShift4Bits
                            
                            If (SurrogateCheck < 1) Or (SurrogateCheck > 16) Then
                                GoTo BadLongCode
                            ElseIf ((b And &HC0) <> &H80) Then
                                GoTo BadLongCode
                            End If
                            
                            CharC = (CharC * vbShift6Bits) Or (b And &H3F)
                            b = Bytes(Index + 1)
                            
                            If (b And &HC0) <> &H80 Then
                                GoTo BadLongCode
                            End If
                            
                            Index = Index + 2
                            CharCount = CharCount + 1
                            MaxFastIndex = MaxFastIndex - 1
                        Else
                            ' 3-byte encoding
                            b = Bytes(Index)
                            
                            If (CharC And (&H1F * vbShift5Bits)) = 0 Then
                                GoTo BadLongCode
                            ElseIf (CharC And (&HF800 \ vbShift6Bits)) = (&HD800& \ vbShift6Bits) Then
                                GoTo BadLongCode
                            ElseIf (b And &HC0) <> &H80 Then
                                GoTo BadLongCode
                            End If
                            
                            Index = Index + 1
                            b = (CharC * vbShift6Bits) Or (b And &H3F)
                            MaxFastIndex = MaxFastIndex - 1
                        End If
                    Else
                        ' 2-byte encoding
                        b = b And &H1F
                        
                        If b <= 1 Then
                            GoTo BadLongCode
                        End If
                    End If
                    
                    CharCount = CharCount + 1
                    MaxFastIndex = MaxFastIndex - 1
                    
ContinueFastLoop:
                Loop
                
                GoTo Continue
                
BadLongCode:
                Index = Index - 2
                GoTo Continue
            End If
            
ProcessChar:
            ' The start of all multi-byte sequence has the 6-bit set. So check
            ' that here for a valid multi-byte sequence.
            If (b And &HE0) = &HC0 Then
                ' 2-byte sequence
                Char = b And &H1F
                
                If Char < 2 Then
                    CharCount = CharCount + CountFallbackCharacter(b, Index, Decoder, Fallback)
                    Char = 0
                Else
                    Char = Char Or Final2ndByte
                End If
            ElseIf (b And &HF0) = &HE0 Then
                ' 3-byte sequence
                Char = (b And &HF) Or ThreeByteStartSeq
            ElseIf (b And &HF8) = &HF0 Then
                ' 4-byte sequence (Surrogate pair)
                Char = b And &H7 'Or FourByteStartSeq
                
                If Char > 4 Then
                    CharCount = CharCount + CountFallbackCharacter(b, Index, Decoder, Fallback)
                    Char = 0
                Else
                    Char = Char Or FourByteStartSeq
                End If
            Else
                ' We have an invalid start for a multi-byte sequence so fallback this byte.
                CharCount = CharCount + CountFallbackCharacter(b, Index, Decoder, Fallback)
            End If
            
            GoTo Continue
        End If
        
        If (b And &HC0) = &H80 Then
            ' Next byte in multi-byte character.
            Char = (Char And &H1FFFFFF) * vbShift6Bits Or (b And &H3F)
            
            ' Check if our initial sequence flag has
            ' reached the end indicating the character is finished.
            If (Char And FinalByte) = 0 Then
                If Char And SupplimentarySeq Then
                    ' It's a 4-byte supplimentary sequence
                    If (Char And Final2ndByte) = 0 Then
                        ' 2nd byte of 4 bytes
                        ' Check for non-shortest form of surrogate and the valid surrogate
                        ' range &h000000 - &h10FFFF at the same time
                        Select Case (Char And &H1F0) - &H10
                            Case Is < 0, Is > 240
                                CharCount = CharCount + CountFallbackCharacter(Char, Index, Decoder, Fallback)
                                Char = 0
                        End Select
                    End If
                Else
                    ' Must be 2nd byte of 3-byte sequence
                    ' Check for non-shortest form of 3 byte sequence
                    If ((Char And &H3E0) = 0) Or ((Char And &H3E0) = &H360) Then
                        CharCount = CharCount + CountFallbackCharacter(Char, Index, Decoder, Fallback)
                        Char = 0
                    End If
                End If
            Else
                If (Char And (SupplimentarySeq Or &H1F0000)) > SupplimentarySeq Then
                    CharCount = CharCount + 1
                    Char = (Char And &H3FF&) + LowSurrogateStart
                End If
            
                Char = 0
                CharCount = CharCount + 1
            End If
        Else
            ' We have an invalid byte for a multi-byte sequence so fallback the character and retry the byte.
            ' Fallback byte index one position to skip the current byte that is to be retried.
            Index = Index - 1
            CharCount = CharCount + CountFallbackCharacter(Char, Index, Decoder, Fallback)
            Char = 0
        End If

Continue:
    Loop
    
    If Decoder Is Nothing Then
        If Char <> 0 Then
            ' Incomplete sequence so fallback character
            CharCount = CharCount + CountFallbackCharacter(Char, Index, Decoder, Fallback)
        End If
    ElseIf Decoder.Flushing Then
        If Char <> 0 Then
            ' Incomplete sequence so fallback character
            CharCount = CharCount + CountFallbackCharacter(Char, Index, Decoder, Fallback)
        End If
    End If
    
    GetCharCountCore = CharCount
End Function

Private Function CountFallbackCharacter(ByVal Char As Long, ByVal ByteIndex As Long, ByVal Decoder As UTF8Decoder, ByRef FallbackBuffer As DecoderFallbackBuffer) As Long
    Select Case SetBytesUnknown(Char)
        Case 1
            CountFallbackCharacter = CountFallbackCharacters(m1ByteUnknown, Char, ByteIndex, Decoder, FallbackBuffer)
        Case 2
            CountFallbackCharacter = CountFallbackCharacters(m2BytesUnknown, Char, ByteIndex, Decoder, FallbackBuffer)
        Case 3
            CountFallbackCharacter = CountFallbackCharacters(m3BytesUnknown, Char, ByteIndex, Decoder, FallbackBuffer)
        Case Else
            Throw Cor.NewArgumentException("Invalid 'BytesUnknown' size.")
    End Select
End Function

Private Function CountFallbackCharacters(ByRef BytesUnknown() As Byte, ByVal Char As Long, ByVal ByteIndex As Long, ByVal Decoder As UTF8Decoder, ByRef FallbackBuffer As DecoderFallbackBuffer) As Long
    If FallbackBuffer Is Nothing Then
        If Decoder Is Nothing Then
            Set FallbackBuffer = mDecoderFallback.CreateFallbackBuffer
        Else
            Set FallbackBuffer = Decoder.FallbackBuffer
        End If
    End If

    FallbackBuffer.Reset
    
    If FallbackBuffer.Fallback(BytesUnknown, ByteIndex - Len1D(BytesUnknown)) Then
        CountFallbackCharacters = FallbackBuffer.Remaining
    End If
End Function

Friend Function GetCharsCore(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal Decoder As UTF8Decoder) As Long
    Dim MaxByteIndex    As Long
    Dim StartCharIndex  As Long
    Dim MaxCharIndex    As Long
    Dim b               As Long
    Dim Char            As Long
    Dim Fallback        As DecoderFallbackBuffer
    Dim CharNoFlags     As Long
    Dim FastBytes()     As Long
    Dim FastChars()     As Long
    Dim FastByteIndex   As Long
    Dim FastCharIndex   As Long
    Dim CanFastLoop     As Boolean
    Dim BytesAddress    As Long
    Dim BytesLB         As Long
    Dim CharsAddress    As Long
    Dim CharsLB         As Long
    
    StartCharIndex = CharIndex
    MaxByteIndex = ByteIndex + ByteCount - 1
    MaxCharIndex = UBound(Chars)
    CanFastLoop = Len1D(Chars) > 1 And Len1D(Bytes) > 3
    
    If CanFastLoop Then
        BytesLB = LBound(Bytes)
        BytesAddress = VarPtr(Bytes(BytesLB))
        CharsLB = LBound(Chars)
        CharsAddress = VarPtr(Chars(CharsLB))
    End If
    
    If Not Decoder Is Nothing Then
        Char = Decoder.Char
    End If
    
    Do While ByteIndex <= MaxByteIndex
        b = Bytes(ByteIndex)
        ByteIndex = ByteIndex + 1

        If Char = 0 Then
            ' Check if we have an ASCII character we can just stuff immediately.
            If b < &H80 Then
                If CharIndex > MaxCharIndex Then _
                    ThrowDecoderSmallConversionBuffer

                Chars(CharIndex) = b
                CharIndex = CharIndex + 1
                
                ' Fast loop: shamelessly adapted from .NET reference source.
                Dim AvailableChars As Long
                Dim AvailableBytes As Long
                Dim MaxFastCharIndex As Long
                
                If Not CanFastLoop Then
                    GoTo Continue
                End If
                
                AvailableChars = MaxCharIndex - CharIndex + 1
                AvailableBytes = MaxByteIndex - ByteIndex + 1
                
                ' Don't fall into the fast decoding loop if we don't have enough characters.
                ' Note that if we don't have enough bytes, MaxFastByteIndex will prevent us from entering the fast loop.
                If AvailableChars <= 13 Then
                    If AvailableChars < AvailableBytes Then
                        GoTo Continue
                    End If
                    
                    ' Run through as many ASCII characters as quickly as we can here.
                    Do While ByteIndex <= MaxByteIndex
                        b = Bytes(ByteIndex)
                        ByteIndex = ByteIndex + 1
                        
                        ' Check if not an ASCII character.
                        If b > &H7F Then
                            GoTo ProcessChar
                        End If
                        
                        Chars(CharIndex) = b
                        CharIndex = CharIndex + 1
                    Loop
                    
                    ' We got here without hitting a non-ascii character so
                    ' we are out of characters. Time to exit the loop.
                    Exit Do
                End If
                
                ' We may need as many as 1 character per byte, so reduce the byte count if necessary.
                ' If AvailableChars is too small, MaxFastByteIndex will be before MaxCharIndex and we won't do fast loop.
                If AvailableChars < AvailableBytes Then
                    AvailableBytes = AvailableChars
                End If
                
                ' To compute the upper bound, assume that all characters are ASCII characters at this point,
                ' the boundary will be decreased for every non-ASCII character we encounter
                ' Also, we need 7 chars reserve for the unrolled ASCII decoding loop and for decoding of multibyte sequences
                MaxFastCharIndex = CharIndex + AvailableBytes - 7
                
                Do While CharIndex < MaxFastCharIndex
                    b = Bytes(ByteIndex)
                    ByteIndex = ByteIndex + 1
                    
                    ' Is it a non-ASCII character?
                    If b > &H7F Then
                        GoTo LongCode
                    End If
                    
                    Chars(CharIndex) = b
                    CharIndex = CharIndex + 1
                    
                    ' Set up our faster access to the Bytes and Chars arrays.
                    mFastLongsSA.PVData = BytesAddress + ByteIndex - BytesLB
                    SAPtr(FastBytes) = mFastLongsSAPtr
                    
                    mFastLongs2SA.PVData = CharsAddress + (CharIndex - CharsLB) * 2
                    SAPtr(FastChars) = mFastLongs2SAPtr
                    
                    FastByteIndex = 0
                    FastCharIndex = 0
                    
                    Do While CharIndex < MaxFastCharIndex
                        Dim b2 As Long
                        b = FastBytes(FastByteIndex)
                        b2 = FastBytes(FastByteIndex + 1)
                        FastByteIndex = FastByteIndex + 2
                        
                        ' Check if any of the 8 bytes is a non-ASCII and bail if any are.
                        If (b Or b2) And &H80808080 Then
                            SAPtr(FastBytes) = vbNullPtr
                            SAPtr(FastChars) = vbNullPtr
                            GoTo LongCodeWithMask
                        End If
                        
                        FastChars(FastCharIndex) = (b And &H7F) Or ((b And &H7F00&) * vbShift8Bits)
                        FastChars(FastCharIndex + 1) = ((b And &H7F0000) \ vbShift16Bits) Or ((b And &H7F000000) \ vbShift8Bits)
                        FastChars(FastCharIndex + 2) = (b2 And &H7F) Or ((b2 And &H7F00&) * vbShift8Bits)
                        FastChars(FastCharIndex + 3) = ((b2 And &H7F0000) \ vbShift16Bits) Or ((b2 And &H7F000000) \ vbShift8Bits)
                        FastCharIndex = FastCharIndex + 4
                        
                        CharIndex = CharIndex + 8
                        ByteIndex = ByteIndex + 8
                    Loop
                    
                    SAPtr(FastBytes) = vbNullPtr
                    SAPtr(FastChars) = vbNullPtr
                    Exit Do

LongCodeWithMask:
                    ' The multi-byte checks in previous code revealed that at least
                    ' one of the bytes is not ASCII, so here we are. Will do a quick
                    ' check to see if the first byte is at least ASCII and if so we
                    ' can store it and continue with the loop. This way there is a
                    ' chance that the multi-byte access didn't go to complete waste.
                    b = b And &HFF&
                    ByteIndex = ByteIndex + 1
                                        
                    If b <= &H7F Then
                        Chars(CharIndex) = b
                        CharIndex = CharIndex + 1
                        GoTo ContinueFastLoop
                    End If
                
LongCode:
                    ' We grab the 2nd byte for the multi-byte sequence.
                    Dim CharC As Long
                    CharC = Bytes(ByteIndex)
                    ByteIndex = ByteIndex + 1
                    
                    ' We do a check to ensure the 0x11XXvvvv 0x10vvvvvv multi-byte format.
                    
                    ' The first byte has already had bit 7 checked earlier, so just check
                    ' bit 6 for the 0x11XXvvvv format. We will know we have at least a 2 byte sequence.
                    '                  ^
                    If (b And &H40) = 0 Then
                        GoTo BadLongCode
                    ' The first byte passed so check the second byte for 0x10vvvvvv format.
                    ElseIf (CharC And &HC0) <> &H80 Then
                        GoTo BadLongCode
                    End If
                    
                    ' Take the encoded lower 6 bits of the second byte in the sequence.
                    CharC = CharC And &H3F
                    
                    ' We check bit 5 to see if this is the start of a 3 or 4 byte sequence.
                    ' 3 byte start with 0x1110vvvv
                    ' 4 byte start with 0x11110vvv
                    '                       ^
                    If b And &H20 Then
                        ' Push the lower 4 bits of the first byte into the
                        ' decoded character sequence, so now we have 9 or 10 bits decoded.
                        CharC = CharC Or ((b And &HF) * vbShift6Bits)
                        
                        ' Check for 4 byte sequence 0x11110vvv.
                        '                                ^
                        If b And &H10 Then
                            ' it's a 4-byte encoding (surrogate pair)
                            b = Bytes(ByteIndex) ' Get byte 3 of the 4-byte sequence.
                            
                            Dim SurrogateCheck As Long
                            ' check that bit 4 is zero, the non-shortest form of surrogate
                            ' and the valid surrogate range 0x000000 - 0x10FFFF at the same time
                            SurrogateCheck = CharC \ vbShift4Bits

                            If (SurrogateCheck < 1) Or (SurrogateCheck > 16) Then
                                GoTo BadLongCode
                            ElseIf (b And &HC0) <> &H80 Then ' byte 3 should have the 0x10vvvvvv format.
                                GoTo BadLongCode
                            End If
                            
                            CharC = (CharC * vbShift6Bits) Or (b And &H3F) ' fold byte 3 bits into our decoded character.
                            b = Bytes(ByteIndex + 1) ' get byte 4 of the sequence.
                            
                            If (b And &HC0) <> &H80 Then ' check that it has the 0x10vvvvvv format.
                                GoTo BadLongCode
                            End If
                            
                            ByteIndex = ByteIndex + 2
                            b = (CharC * vbShift6Bits) Or (b And &H3F) ' fold byte 4 bits into the decoded character.
                            
                            ' we should have a value from 0x000000 to 0x10FFFF.
                            ' Push the high surrogate part into a character and let the low surrogate character get
                            ' stored at the end of the multi-byte sequence.
#If Release Then
                            Chars(CharIndex) = (((b \ vbShift10Bits) And &H7FF&) + (HighSurrogateStart - 64)) And &HFFFF&
#Else
                            Chars(CharIndex) = AsWord(((b \ vbShift10Bits) And &H7FF&) + (HighSurrogateStart - 64))
#End If
                            CharIndex = CharIndex + 1
                            
                            b = (b And &H3FF) + LowSurrogateStart
                            MaxFastCharIndex = MaxFastCharIndex - 1
                        Else
                            ' must be a 3-byte encoding
                            b = Bytes(ByteIndex)
                            
                            '
                            If (CharC And &H3E0) = 0 Then
                                GoTo BadLongCode
                            ElseIf (CharC And &H3E0) = &H360 Then
                                GoTo BadLongCode
                            ElseIf (b And &HC0) <> &H80 Then
                                GoTo BadLongCode
                            End If
                            
                            ByteIndex = ByteIndex + 1
                            b = (CharC * vbShift6Bits) Or (b And &H3F)
                            MaxFastCharIndex = MaxFastCharIndex - 1
                        End If
                    Else
                        ' 2-byte encoding
                        b = b And &H1F
                        
                        If b <= 1 Then
                            GoTo BadLongCode
                        End If
                        
                        b = (b * vbShift6Bits) Or CharC
                    End If
                    
#If Release Then
                    Chars(CharIndex) = (b And &HFFFF&)
#Else
                    Chars(CharIndex) = AsWord(b)
#End If
                    CharIndex = CharIndex + 1
                    MaxFastCharIndex = MaxFastCharIndex - 1
                    
ContinueFastLoop:
                Loop
                
                GoTo Continue
                
BadLongCode:
                ' Move back 2 positions because we only officially move
                ' the index forward after validation has occurred. So only
                ' the first 2 bytes of a multi-byte sequence has actually been counted for.
                ByteIndex = ByteIndex - 2
                GoTo Continue
            End If
            
ProcessChar:
            ' The start of all multi-byte sequence has the 6-bit set. So check
            ' that here for a valid multi-byte sequence.
            If (b And &HE0) = &HC0 Then
                ' 2-byte sequence
                Char = b And &H1F
                
                If Char < 2 Then
                    CharIndex = FallbackCharacter(b, ByteIndex, Chars, CharIndex, Decoder, Fallback)
                    Char = 0
                Else
                    Char = Char Or Final2ndByte
                End If
            ElseIf (b And &HF0) = &HE0 Then
                ' 3-byte sequence
                Char = (b And &HF) Or ThreeByteStartSeq
            ElseIf (b And &HF8) = &HF0 Then
                ' 4-byte sequence (Surrogate pair)
                Char = b And &H7 'Or FourByteStartSeq
                
                If Char > 4 Then
                    CharIndex = FallbackCharacter(b Or &HF0, ByteIndex, Chars, CharIndex, Decoder, Fallback)
                    Char = 0
                Else
                    Char = Char Or FourByteStartSeq
                End If
            Else
                ' We have an invalid start for a multi-byte sequence so fallback this byte.
                CharIndex = FallbackCharacter(b, ByteIndex, Chars, CharIndex, Decoder, Fallback)
            End If
            
            GoTo Continue
        End If
        
        
        If (b And &HC0) = &H80 Then
            ' Next byte in multi-byte character.
            Char = (Char And &H1FFFFFF) * vbShift6Bits Or (b And &H3F)
            
            ' Check if our initial sequence flag has
            ' reached the end indicating the character is finished.
            If (Char And FinalByte) = 0 Then
                If Char And SupplimentarySeq Then
                    ' It's a 4-byte supplimentary sequence
                    If (Char And Final2ndByte) = 0 Then
                        ' 2nd byte of 4 bytes
                        ' Check for non-shortest form of surrogate and the valid surrogate
                        ' range &h000000 - &h10FFFF at the same time
                        Select Case Char And &H1F0
                            Case Is < &H10, Is > &H100
                                CharIndex = FallbackCharacter(Char, ByteIndex, Chars, CharIndex, Decoder, Fallback)
                                Char = 0
                        End Select
                    End If
                Else
                    ' Must be 2nd byte of 3-byte sequence
                    ' Check for non-shortest form of 3 byte sequence
                    If ((Char And &H3E0) = 0) Or ((Char And &H3E0) = &H360) Then
                        CharIndex = FallbackCharacter(Char, ByteIndex, Chars, CharIndex, Decoder, Fallback)
                        Char = 0
                    End If
                End If
            Else
                If (Char And (SupplimentarySeq Or &H1F0000)) > SupplimentarySeq Then
                    If CharIndex <= MaxCharIndex Then
#If Release Then
                        Chars(CharIndex) = (((Char \ vbShift10Bits) And &H7FF&) + HighSurrogateStart - &H40) And &HFFFF&
#Else
                        Chars(CharIndex) = AsWord(((Char \ vbShift10Bits) And &H7FF&) + HighSurrogateStart - &H40)
#End If
                        CharIndex = CharIndex + 1
                        Char = (Char And &H3FF&) + LowSurrogateStart
                    End If
                End If
            
                If CharIndex > MaxCharIndex Then _
                        ThrowDecoderSmallConversionBuffer
                                    
#If Release Then
                    Chars(CharIndex) = Char And &HFFFF&
#Else
                    Chars(CharIndex) = AsWord(Char)
#End If
                Char = 0
                CharIndex = CharIndex + 1
            End If
        Else
            ' We have an invalid byte for a multi-byte sequence so fallback the character and retry the byte.
            ' Fallback byte index one position to skip the current byte that is to be retried.
            ByteIndex = ByteIndex - 1
            CharIndex = FallbackCharacter(Char, ByteIndex, Chars, CharIndex, Decoder, Fallback)
            Char = 0
        End If

Continue:
    Loop
    
    If Decoder Is Nothing Then
        If Char <> 0 Then
            ' Incomplete sequence so fallback character
            CharIndex = FallbackCharacter(Char, ByteIndex, Chars, CharIndex, Decoder, Fallback)
        End If
    ElseIf Decoder.Flushing Then
        If Char <> 0 Then
            ' Incomplete sequence so fallback character
            CharIndex = FallbackCharacter(Char, ByteIndex, Chars, CharIndex, Decoder, Fallback)
            Decoder.Char = 0
        End If
    Else
        Decoder.Char = Char
    End If
    
    GetCharsCore = CharIndex - StartCharIndex
End Function

Private Function FallbackCharacter(ByVal Char As Long, ByVal ByteIndex As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal Decoder As UTF8Decoder, ByRef FallbackBuffer As DecoderFallbackBuffer) As Long
    Select Case SetBytesUnknown(Char)
        Case 1
            FallbackCharacter = InsertFallbackCharacters(m1ByteUnknown, Char, ByteIndex, Chars, CharIndex, Decoder, FallbackBuffer)
        Case 2
            FallbackCharacter = InsertFallbackCharacters(m2BytesUnknown, Char, ByteIndex, Chars, CharIndex, Decoder, FallbackBuffer)
        Case 3
            FallbackCharacter = InsertFallbackCharacters(m3BytesUnknown, Char, ByteIndex, Chars, CharIndex, Decoder, FallbackBuffer)
        Case Else
            Throw Cor.NewArgumentException("Invalid 'BytesUnknown' size.")
    End Select
End Function

Private Function InsertFallbackCharacters(ByRef BytesUnknown() As Byte, ByVal Char As Long, ByVal ByteIndex As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal Decoder As UTF8Decoder, ByRef FallbackBuffer As DecoderFallbackBuffer) As Long
    Dim MaxCharIndex As Long
    
    If FallbackBuffer Is Nothing Then
        If Decoder Is Nothing Then
            Set FallbackBuffer = mDecoderFallback.CreateFallbackBuffer
        Else
            Set FallbackBuffer = Decoder.FallbackBuffer
        End If
    End If

    FallbackBuffer.Reset
    
    ' The .NET version of the fallback will have a DecoderFallbackException.Index of one less than
    ' what the documentation says. The documentation says the index is of the byte that caused
    ' the invalid sequence to happen. I'm going with the documentation because that index value
    ' makes the most sense, so all VBCorLib indexes will be one more than what .NET indicates.
    If FallbackBuffer.Fallback(BytesUnknown, ByteIndex - Len1D(BytesUnknown)) Then
        MaxCharIndex = UBound(Chars)
        
        Do While FallbackBuffer.Remaining > 0
            Dim Ch As Integer
            Ch = FallbackBuffer.GetNextChar
            
            If CharIndex > MaxCharIndex Then _
                ThrowDecoderSmallConversionBuffer
                
            Chars(CharIndex) = Ch
            CharIndex = CharIndex + 1
        Loop
    End If
    
    InsertFallbackCharacters = CharIndex
End Function

Private Function SetBytesUnknown(ByVal Ch As Long) As Long
    Dim Size As Long
    
    ' See if it was a plain char
    If Ch > 0 And Ch < &H100 Then
        m1ByteUnknown(0) = Ch
        Size = 1
    ' See if its an unfinished 2 byte sequence
    ElseIf (Ch And (SupplimentarySeq Or ThreeByteSeq)) = 0 Then
        m1ByteUnknown(0) = (Ch And &H1F) Or &HC0
        Size = 1
    ' So now we're either 2nd byte of 3 or 4 byte sequence or
    ' we hit a non-trail byte or we ran out of space for 3rd byte of 4 byte sequence
    ' 1st check if its a 4 byte sequence
    ElseIf Ch And SupplimentarySeq Then
        If Ch And Final2ndByte Then
            ' 3rd byte of 4 byte sequence
            m3BytesUnknown(0) = ((Ch \ vbShift12Bits) And &H7) Or &HF0
            m3BytesUnknown(1) = ((Ch \ vbShift6Bits) And &H3F) Or &H80
            m3BytesUnknown(2) = (Ch And &H3F) Or &H80
            Size = 3
        ElseIf Ch And Final3rdByte Then
            ' 2nd byte of a 4 byte sequence
            m2BytesUnknown(0) = ((Ch \ vbShift6Bits) And &H7) Or &HF0
            m2BytesUnknown(1) = (Ch And &H3F) Or &H80
            Size = 2
        Else
            ' 4th byte of a 4 byte sequence
            m1ByteUnknown(0) = (Ch And &H7) Or &HF0
            Size = 1
        End If
    Else
        ' 2nd byte of 3 byte sequence?
        If Ch And Final2ndByte Then
            m2BytesUnknown(0) = ((Ch \ vbShift6Bits) And &HF) Or &HE0
            m2BytesUnknown(1) = (Ch And &H3F) Or &H80
            Size = 2
        Else
            ' 1st byte of a 3 byte sequence
            m1ByteUnknown(0) = (Ch And &HF) Or &HE0
            Size = 1
        End If
    End If
    
    SetBytesUnknown = Size
End Function

Friend Function GetByteCountCore(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByVal Encoder As UTF8Encoder) As Long
    Dim Ch              As Long
    Dim PrevChar        As Long
    Dim Fallback        As EncoderFallbackBuffer
    Dim InFallback      As Boolean
    Dim MaxCharIndex    As Long
    Dim StartByteIndex  As Long
    Dim FastChars()     As Long
    Dim CanFastLoop     As Boolean
    Dim CharsAddress    As Long
    Dim ByteCount       As Long
    Dim CharsLB         As Long
    
    MaxCharIndex = Index + Count - 1
    CanFastLoop = Len1D(Chars) > 0
    
    If CanFastLoop Then
        CharsLB = LBound(Chars)
        CharsAddress = VarPtr(Chars(CharsLB))
    End If
    
    If Not Encoder Is Nothing Then
        PrevChar = Encoder.PrevChar
    End If
    
ProcessNext:
    Do While Index <= MaxCharIndex Or InFallback
        If InFallback Then
            If Fallback.Remaining > 0 Then
                Ch = Fallback.GetNextChar And &HFFFF&
            Else
                InFallback = False
                GoTo Continue
            End If
        Else
            Ch = Chars(Index) And &HFFFF&
            Index = Index + 1
            
            ' Quick check to see if we should even attempt a fast loop because
            ' we won't know for sure until everything has been set up then
            ' it needs to be torn down. This causes huge slow-downs for non-ascii characters.
            If Ch < &H80 Then
                If Not InFallback And PrevChar = 0 And CanFastLoop Then
                    ' Getting here means we will at least count the initial char.
                    ByteCount = ByteCount + 1
                    
                    Dim MaxFastIndex    As Long
                    Dim FastIndex       As Long
                    
                    MaxFastIndex = (Count - Index) \ 2 - 1
                    mFastLongsSA.PVData = CharsAddress + (Index - CharsLB) * 2
                    SAPtr(FastChars) = mFastLongsSAPtr
                    
                    For FastIndex = 0 To MaxFastIndex
                        If FastChars(FastIndex) And &HFF80FF80 Then
                            Exit For
                        End If
                    Next
                    
                    SAPtr(FastChars) = vbNullPtr
                    
                    If FastIndex > 0 Then
                        Index = Index + FastIndex * 2
                        ByteCount = ByteCount + FastIndex * 2
                    End If
                    
                    GoTo Continue
                End If
            End If
        End If
        
ProcessChar:
        If PrevChar = 0 Then
            Select Case Ch
                Case Is < &H80
                    ByteCount = ByteCount + 1
                Case Is < &H800
                    ByteCount = ByteCount + 2
                Case HighSurrogateStart To HighSurrogateEnd
                    PrevChar = Ch
                Case LowSurrogateStart To LowSurrogateEnd
                    ' a low surrogate without a preceeding high surrogate
                    ' just fallback the low surrogate.
                    InFallback = FallbackUnknownChar(Ch, Index - 1, Fallback, Encoder)
                Case Else   ' 0800-FFFF excluding surrogates; 3-byte UTF8 encoded
                    ByteCount = ByteCount + 3
            End Select
        Else
            Select Case Ch
                Case LowSurrogateStart To LowSurrogateEnd
                    ' we have a surrogate pair
                    ByteCount = ByteCount + 4
                Case Else
                    ' we have a high surrogate with no following low surrogate
                    ' so fallback the high surrogate. move the index back so
                    ' we will attempt to re-process the current character once
                    ' the fallback has completed.
                    Index = Index - 1
                    InFallback = FallbackUnknownChar(PrevChar, Index - 1, Fallback, Encoder)
            End Select
    
            PrevChar = 0
        End If
        
Continue:
    Loop
    
    If Encoder Is Nothing Then
        If PrevChar <> 0 Then
            ' Fallback hanging high-surrogate
            InFallback = FallbackUnknownChar(PrevChar, Index - 1, Fallback, Encoder)
            PrevChar = 0
            GoTo ProcessNext
        End If
    ElseIf Encoder.Flush Then
        If PrevChar <> 0 Then
            ' Fallback hanging high-surrogate
            InFallback = FallbackUnknownChar(PrevChar, Index - 1, Fallback, Encoder)
            PrevChar = 0
            GoTo ProcessNext
        End If
    End If
    
    GetByteCountCore = ByteCount
End Function

Friend Function GetBytesCore(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal Encoder As UTF8Encoder) As Long
    Dim Ch              As Long
    Dim PrevChar        As Long
    Dim Fallback        As EncoderFallbackBuffer
    Dim InFallback      As Boolean
    Dim MaxCharIndex    As Long
    Dim StartByteIndex  As Long
    Dim MaxByteIndex    As Long
    Dim FastChars()     As Long
    Dim FastBytes()     As Long
    Dim CanFastLoop     As Boolean
    Dim CharsAddress    As Long
    Dim BytesAddress    As Long
    Dim CharsLB         As Long
    Dim BytesLB         As Long
    
    StartByteIndex = ByteIndex
    MaxByteIndex = UBound(Bytes)
    MaxCharIndex = CharIndex + CharCount - 1
    CanFastLoop = Len1D(Chars) > 0 And Len1D(Bytes) > 0
    
    If CanFastLoop Then
        CharsLB = LBound(Chars)
        BytesLB = LBound(Bytes)
        CharsAddress = VarPtr(Chars(CharsLB))
        BytesAddress = VarPtr(Bytes(BytesLB))
    End If
    
    If Not Encoder Is Nothing Then
        PrevChar = Encoder.PrevChar
        Encoder.PrevChar = 0
    End If
    
ProcessNext:
    Do While CharIndex <= MaxCharIndex Or InFallback
        If InFallback Then
            If Fallback.Remaining > 0 Then
                Ch = Fallback.GetNextChar And &HFFFF&
            Else
                InFallback = False
                GoTo Continue
            End If
        Else
            Ch = Chars(CharIndex) And &HFFFF&
            CharIndex = CharIndex + 1
        End If
        
        If PrevChar = 0 Then
            If Ch < &H80 Then
                If ByteIndex > MaxByteIndex Then _
                    ThrowEncoderSmallConversionBuffer
                
                Bytes(ByteIndex) = Ch
                ByteIndex = ByteIndex + 1
                
                If Not CanFastLoop Then
                    GoTo Continue
                End If
                
                ' Fast loop: shamelessly adapted from the .NET reference source.
                Dim AvailableChars  As Long
                Dim AvailableBytes  As Long
                Dim MaxFastCharIndex As Long
                Dim FastCharIndex   As Long
                Dim FastByteIndex   As Long
                Dim ChB             As Long
                
                AvailableChars = MaxCharIndex - CharIndex + 1
                AvailableBytes = MaxByteIndex - ByteIndex + 1

                ' Don't fall into the fast decoding loop if we don't have enough characters
                ' Note that if we don't have enough bytes, MaxFastCharIndex will prevent us from entering the fast loop.
                If AvailableChars <= 13 Then
                    ' we are hoping for 1 byte per char.
                    If AvailableBytes < AvailableChars Then
                        ' Not enough output room.
                        GoTo Continue
                    End If

                    ' Run through the characters as fast as we can in
                    ' a tight loop as long as they are ascii.
                    Do While CharIndex <= MaxCharIndex
                        Ch = Chars(CharIndex) And &HFFFF&
                        CharIndex = CharIndex + 1

                        If Ch > &H7F Then
                            ' we hit a non-ascii character so process it through the slow loop.
                            GoTo ProcessNextChar
                        End If

                        Bytes(ByteIndex) = Ch
                        ByteIndex = ByteIndex + 1
                    Loop

                    ' Since we got here we must have looped through all
                    ' characters and stored them in the output bytes.
                    ' So exit the slow loop; we are done.
                    Exit Do
                End If
                
                ' We are going to assume that all chars are ASCII at this point.
                ' Keep two chars in reserve for the ChB access of the chars array.
                If AvailableChars < AvailableBytes Then
                    MaxFastCharIndex = CharIndex + AvailableChars - 3
                Else
                    MaxFastCharIndex = CharIndex + AvailableBytes - 3
                End If
                                                
                ' Prepare for fast access to both the character and byte arrays.
                mFastLongsSA.PVData = CharsAddress + (CharIndex - CharsLB) * 2
                mFastLongs2SA.PVData = BytesAddress + ByteIndex - BytesLB
                SAPtr(FastChars) = mFastLongsSAPtr
                SAPtr(FastBytes) = mFastLongs2SAPtr
                FastCharIndex = 0
                FastByteIndex = 0

                ' Run 4 characters at a time!
                Do While CharIndex < MaxFastCharIndex
                    Ch = FastChars(FastCharIndex)
                    ChB = FastChars(FastCharIndex + 1)
                
                    ' Check for any non-ascii characters.
                    If (Ch Or ChB) And &HFF80FF80 Then
                        SAPtr(FastChars) = vbNullPtr
                        SAPtr(FastBytes) = vbNullPtr
                        GoTo ExitFastLoop
                    End If

                    FastBytes(FastByteIndex) = (Ch And &H7F&) Or ((Ch And &H7F0000) \ vbShift8Bits) Or ((ChB And &H7F&) * vbShift16Bits) Or ((ChB And &H7F0000) * vbShift8Bits)
                    FastByteIndex = FastByteIndex + 1
                    FastCharIndex = FastCharIndex + 2
                    CharIndex = CharIndex + 4
                    ByteIndex = ByteIndex + 4
                Loop
                
                SAPtr(FastChars) = vbNullPtr
                SAPtr(FastBytes) = vbNullPtr
                GoTo Continue
                
ExitFastLoop:
                ' We came from the inner fast loop with the
                ' Ch holding 2 characters, so we need to
                ' chop off the high 16-bits and process
                ' the lower 16-bits as a regulare character.
                Ch = Ch And &HFFFF&

                If Ch <= &H7F Then
                    Bytes(ByteIndex) = Ch
                    CharIndex = CharIndex + 1
                    ByteIndex = ByteIndex + 1
                End If

                GoTo Continue
            End If
            
ProcessNextChar:
            Select Case Ch
                ' We aren't concerned about getting a character less than &h80
                ' here because it would already have been dealt with earlier.
                Case Is < &H800
                    If ByteIndex > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer

                    ' Encode 2-byte character.
                    ' 110vvvvv 10vvvvvv
                    ' ^^^ = &hC0
                    Bytes(ByteIndex) = &HC0 Or (Ch \ vbShift6Bits)
                    Bytes(ByteIndex + 1) = &H80 Or (Ch And &H3F)
                    ByteIndex = ByteIndex + 2
                Case HighSurrogateStart To HighSurrogateEnd
                    PrevChar = Ch
                Case LowSurrogateStart To LowSurrogateEnd
                    ' a low surrogate without a preceeding high surrogate
                    ' just fallback the low surrogate.
                    InFallback = FallbackUnknownChar(Ch, CharIndex - 1, Fallback, Encoder)
                Case Else   ' 0800-FFFF excluding surrogates; 3-byte UTF8 encoded
                    If ByteIndex + 2 > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer

                    Bytes(ByteIndex) = &HE0 Or ((Ch \ &H1000&) And &HF)
                    Bytes(ByteIndex + 1) = &H80 Or ((Ch \ vbShift6Bits) And &H3F)
                    Bytes(ByteIndex + 2) = &H80 Or (Ch And &H3F)
                    ByteIndex = ByteIndex + 3
            End Select
        Else
            Select Case Ch
                Case LowSurrogateStart To LowSurrogateEnd
                    ' we have a surrogate pair
                    Dim UnicodeChar As Long
                    If ByteIndex + 3 > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer

                    UnicodeChar = ((PrevChar - HighSurrogateStart) * &H400) + (Ch - LowSurrogateStart) + &H10000
                    Bytes(ByteIndex + 3) = &H80 Or (UnicodeChar And &H3F)
                    Bytes(ByteIndex + 2) = &H80 Or ((UnicodeChar \ &H40&) And &H3F)
                    Bytes(ByteIndex + 1) = &H80 Or ((UnicodeChar \ &H1000&) And &H3F)
                    Bytes(ByteIndex) = &HF0 Or (UnicodeChar \ &H40000)
                    ByteIndex = ByteIndex + 4
                Case Else
                    ' we have a high surrogate with no following low surrogate
                    ' so fallback the high surrogate. move the index back so
                    ' we will attempt to re-process the current character once
                    ' the fallback has completed.
                    CharIndex = CharIndex - 1
                    InFallback = FallbackUnknownChar(PrevChar, CharIndex - 1, Fallback, Encoder)
            End Select
    
            PrevChar = 0
        End If
        
Continue:
    Loop
    
    If Encoder Is Nothing Then
        If PrevChar <> 0 Then
            ' Fallback hanging high-surrogate
            InFallback = FallbackUnknownChar(PrevChar, CharIndex - 1, Fallback, Encoder)
            PrevChar = 0
            GoTo ProcessNext
        End If
    ElseIf Encoder.Flush Then
        If PrevChar <> 0 Then
            ' Fallback hanging high-surrogate
            InFallback = FallbackUnknownChar(PrevChar, CharIndex - 1, Fallback, Encoder)
            PrevChar = 0
            GoTo ProcessNext
        End If
    Else
        Encoder.PrevChar = PrevChar
    End If
    
    GetBytesCore = ByteIndex - StartByteIndex
End Function

Private Function FallbackUnknownChar(ByVal UnknownChar As Long, ByVal CharIndex As Long, ByRef Fallback As EncoderFallbackBuffer, ByVal Encoder As Encoder) As Boolean
    If Fallback Is Nothing Then
        If Encoder Is Nothing Then
            Set Fallback = mEncoderFallback.CreateFallbackBuffer
        Else
            Set Fallback = Encoder.FallbackBuffer
        End If
    End If
    
    FallbackUnknownChar = Fallback.Fallback(UnknownChar, Index:=CharIndex)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Table() As EncodingTable
    Set Table = Statics.Encoding.EncodingTable
End Property

Private Sub ThrowEncoderSmallConversionBuffer()
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowBytes, EncodingName, TypeName(EncoderFallback)), "Bytes")
End Sub

Private Sub ThrowDecoderSmallConversionBuffer()
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowChars, EncodingName, TypeName(DecoderFallback)), "Chars")
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mShouldEmitUTF8Identifier = .ReadProperty(EmitUTF8IdentifierProp)
        mIsReadOnly = .ReadProperty(IsReadOnlyProp)
        Set mDecoderFallback = .ReadProperty(DecoderFallbackProp, Statics.DecoderFallback.ExceptionFallback)
        Set mEncoderFallback = .ReadProperty(EncoderFallbackProp, Statics.EncoderFallback.ExceptionFallback)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty EmitUTF8IdentifierProp, mShouldEmitUTF8Identifier
        .WriteProperty IsReadOnlyProp, mIsReadOnly
        
        If Not mDecoderFallback.Equals(Statics.DecoderFallback.ExceptionFallback) Then
            .WriteProperty DecoderFallbackProp, mDecoderFallback
        End If
        
        If Not mEncoderFallback.Equals(Statics.EncoderFallback.ExceptionFallback) Then
            .WriteProperty EncoderFallbackProp, mEncoderFallback
        End If
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Encoding
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Encoding_IsReadOnly() As Boolean
    Encoding_IsReadOnly = IsReadOnly
End Property

Private Function Encoding_Clone() As Encoding
    Set Encoding_Clone = Clone
End Function

Private Property Set Encoding_DecoderFallback(ByVal Value As DecoderFallback)
    Set DecoderFallback = Value
End Property

Private Property Get Encoding_DecoderFallback() As DecoderFallback
    Set Encoding_DecoderFallback = DecoderFallback
End Property

Private Property Set Encoding_EncoderFallback(ByVal Value As EncoderFallback)
    Set EncoderFallback = Value
End Property

Private Property Get Encoding_EncoderFallback() As EncoderFallback
    Set Encoding_EncoderFallback = EncoderFallback
End Property

Private Property Get Encoding_BodyName() As String
    Encoding_BodyName = BodyName
End Property

Private Property Get Encoding_CodePage() As Long
    Encoding_CodePage = CodePage
End Property

Private Property Get Encoding_EncodingName() As String
    Encoding_EncodingName = EncodingName
End Property

Private Function Encoding_Equals(Value As Variant) As Boolean
    Encoding_Equals = Equals(Value)
End Function

Private Function Encoding_GetByteCount(Value As Variant, Optional Index As Variant, Optional Count As Variant) As Long
    Encoding_GetByteCount = GetByteCount(Value, Index, Count)
End Function

Private Function Encoding_GetBytes(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    Encoding_GetBytes = GetBytes(Source, Index, Count)
End Function

Private Function Encoding_GetBytesEx(Source As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Encoding_GetBytesEx = GetBytesEx(Source, CharIndex, CharCount, Bytes, ByteIndex)
End Function

Private Function Encoding_GetCharCount(Bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As Long
    Encoding_GetCharCount = GetCharCount(Bytes, Index, Count)
End Function

Private Function Encoding_GetChars(Bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As Integer()
    Encoding_GetChars = GetChars(Bytes, Index, Count)
End Function

Private Function Encoding_GetCharsEx(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Encoding_GetCharsEx = GetCharsEx(Bytes, ByteIndex, ByteCount, Chars, CharIndex)
End Function

Private Function Encoding_GetDecoder() As Decoder
    Set Encoding_GetDecoder = GetDecoder
End Function

Private Function Encoding_GetEncoder() As Encoder
    Set Encoding_GetEncoder = GetEncoder
End Function

Private Function Encoding_GetHashCode() As Long
    Encoding_GetHashCode = GetHashCode
End Function

Private Function Encoding_GetMaxByteCount(ByVal CharCount As Long) As Long
    Encoding_GetMaxByteCount = GetMaxByteCount(CharCount)
End Function

Private Function Encoding_GetMaxCharCount(ByVal ByteCount As Long) As Long
    Encoding_GetMaxCharCount = GetMaxCharCount(ByteCount)
End Function

Private Function Encoding_GetPreamble() As Byte()
    Encoding_GetPreamble = GetPreamble
End Function

Private Function Encoding_GetString(Bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As String
    Encoding_GetString = GetString(Bytes, Index, Count)
End Function

Private Property Get Encoding_HeaderName() As String
    Encoding_HeaderName = HeaderName
End Property

Private Property Get Encoding_IsBrowserDisplay() As Boolean
    Encoding_IsBrowserDisplay = IsBrowserDisplay
End Property

Private Property Get Encoding_IsBrowserSave() As Boolean
    Encoding_IsBrowserSave = IsBrowserSave
End Property

Private Property Get Encoding_IsMailNewsDisplay() As Boolean
    Encoding_IsMailNewsDisplay = IsMailNewsDisplay
End Property

Private Property Get Encoding_IsMailNewsSave() As Boolean
    Encoding_IsMailNewsSave = IsMailNewsSave
End Property

Private Function Encoding_ToString() As String
    Encoding_ToString = ToString
End Function

Private Property Get Encoding_WebName() As String
    Encoding_WebName = WebName
End Property

Private Property Get Encoding_WindowsCodePage() As Long
    Encoding_WindowsCodePage = WindowsCodePage
End Property

Private Property Get Encoding_IsSingleByte() As Boolean
    Encoding_IsSingleByte = IsSingleByte
End Property

