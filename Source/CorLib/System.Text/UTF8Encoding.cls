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

''
' A set of functions to be used to convert character arrays to and from byte arrays.
'
' @remarks <p>UTF-8 is an encoding scheme that is fully compatible with the existing
' ASCII set of characters (0-127). It also supports Unicode characters, including the
' UTF-16 character set.</p>
' <p>UTF-8 is useful because it will have a size ratio of 1:1 for pure ASCII characters.
' The ratio will increase during the encoding for characters outside ASCII.</p>
' <p>Additional information about UTF-8 encoding scheme can be found at "http://www.faqs.org/rfcs/rfc2279.html".</p>
'
' @see Constructors
' @see Encoding
'
Option Explicit
Implements IObject
Implements Encoding

Private Const CODE_PAGE As Long = 65001

Private Const Prop_ThrowException = "ThrowException"
Private Const Prop_ShouldEmitUTF8Identifier = "EmitBOM"

Private mIsReadOnly                 As Boolean
Private mThrowException             As Boolean
Private mShouldEmitUTF8Identifier   As Boolean
Private mEncoderFallback            As EncoderFallback
Private mDecoderFallback            As DecoderFallback
Private mBytesUnknown(0 To 1)       As Byte


''
' When implemented in a derived class, gets a value indicating whether the current encoding is read-only.
'
' @return <b>true</b> if the current <a href="Encoding.html">Encoding</a> is read-only; otherwise, <b>false</b>. The default is <b>true</b>.
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
' Returns if the current encoding uses single-byte code points.
'
' @return Returns True for single-byte, False for multi-byte code points.
'
Public Property Get IsSingleByte() As Boolean
    IsSingleByte = False
End Property

''
' Returns the encoding name to be used in with the mail agent body tags.
'
' @return The body tag compatible encoding name.
'
Public Property Get BodyName() As String
    BodyName = Table.BodyName(CodePage)
End Property

''
' Returns the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
'
Public Property Get WebName() As String
    WebName = Table.WebName(CodePage)
End Property

''
' Returns the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
'
Public Property Get HeaderName() As String
    HeaderName = Table.HeaderName(CodePage)
End Property

''
' Returns the code page identifier for this encoding.
'
' @return Code page identifier.
'
Public Property Get CodePage() As Long
    CodePage = CODE_PAGE
End Property

''
' Returns the Windows Operating System's code page for this encoding.
'
' @return The Windows code page for this encoding.
'
Public Property Get WindowsCodePage() As Long
    WindowsCodePage = Table.WindowsCodePage(CodePage)
End Property

''
' The name of this encoding.
'
' @return The encoding's name.
'
Public Property Get EncodingName() As String
    EncodingName = Table.EncodingName(CodePage)
End Property

''
' Indicates if this encoding can be used by browsers to display text.
'
' @return The indication for browser display capable.
'
Public Property Get IsBrowserDisplay() As Boolean
    IsBrowserDisplay = Table.IsBrowserDisplay(CodePage)
End Property

''
' Indicates if this encoding can be used to save data with this encoding.
'
' @return The indication for the browser saving capable.
'
Public Property Get IsBrowserSave() As Boolean
    IsBrowserSave = Table.IsBrowserSave(CodePage)
End Property

''
' Indicates if this encoding can be used to display mail and news by
' mail and news clients.
'
' @return Indication for mail and news client capable.
'
Public Property Get IsMailNewsDisplay() As Boolean
    IsMailNewsDisplay = Table.IsMailNewsDisplay(CodePage)
End Property

''
' Indicates if this encoding can be used to save data by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
'
Public Property Get IsMailNewsSave() As Boolean
    IsMailNewsSave = Table.IsMailNewsSave(CodePage)
End Property

''
' Returns the current DecoderFallback instance used by the encoding.
'
' @return An object derived from <a href="DecoderFallback.html">DecoderFallback</a>.
' @remarks The default will return an instance of <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a>.
'
Public Property Get DecoderFallback() As DecoderFallback
    Set DecoderFallback = mDecoderFallback
End Property

''
' Sets the <a href="DecoderFallback.html">DecoderFallback</a> to be used by this encoding instance.
'
' @param Value The <a href="DecoderFallback.html">DecoderFallback</a> instance to set.
' @remarks By default an <a href="Encoding">Encoding</a> instance is marked as read-only and setting
' this value will result in an <a href="InvalidOperationException.html">InvalidOperationException</a>
' being thrown. If a clone of the instance is made, then it will become read-write and allow the
' setting of the value.
'
Public Property Set DecoderFallback(ByVal Value As DecoderFallback)
    If mIsReadOnly Then _
        Error.InvalidOperation InvalidOperation_ReadOnly
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    
    Set mDecoderFallback = Value
End Property

''
' Returns the current EncoderFallback instance used by the encoding.
'
' @return An object derived from <a href="EncoderFallback.html">EncoderFallback</a>.
' @remarks The default will return an instance of <a href="EncoderReplacementFallback.html">EncoderReplacementFallback</a>.
'
Public Property Get EncoderFallback() As EncoderFallback
    Set EncoderFallback = mEncoderFallback
End Property

''
' Sets the <a href="EncoderFallback.html">EncoderFallback</a> to be used by this encoding instance.
'
' @param Value The <a href="EncoderFallback.html">EncoderFallback</a> instance to set.
' @remarks By default an <a href="Encoding">Encoding</a> instance is marked as read-only and setting
' this value will result in an <a href="InvalidOperationException.html">InvalidOperationException</a>
' being thrown. If a clone of the instance is made, then it will become read-write and allow the
' setting of the value.
'
Public Property Set EncoderFallback(ByVal Value As EncoderFallback)
    If mIsReadOnly Then _
        Error.InvalidOperation InvalidOperation_ReadOnly
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    
    Set mEncoderFallback = Value
End Property

''
' Returns the number of bytes that would be produced from the set of characters using this encoding.
'
' @param Chars A set of characters used to calculated the number of bytes once encoded with this encoding.
' @param Index The starting index of the first character to be used.
' @param Count The number of characters to include in the calculation.
' @return The number of bytes that would be produced by encoding the set of characters.
' @remarks The <i>Chars</i> parameter will accept either a String or an Integer array. The Index
' is zero based for <b>String</b>, not 1 based like Mid$.
'
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim NullBytes()     As Byte
    Dim LocalChars()    As Integer
    Dim Range           As ListRange
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    Range = ValidateGetBytes(LocalChars, Index, Count)
    
    If Range.Count > 0 Then
        GetByteCount = GetBytesCore(LocalChars, Range.Index, Range.Count, NullBytes, 0, Nothing)
    End If
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    FreeChars LocalChars
    Return
End Function

''
' Encodes a set of characters into an array of bytes.
'
' @param Chars The set of characters to be encoded into an array of bytes. This can
' be either an Integer array or a String.
' @param Index The index of the first character to begin encoding from.
' @param Count The number of characters to be encoded.
' @return A byte array containing the encoded characters as bytes.
' @remarks The <i>Chars</i> parameter will accept either a String or an Integer array. The Index
' is zero based for both the String and Integer array, not 1 based like Mid$.
'
Public Function GetBytes(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    Dim Range           As ListRange
    Dim LocalChars()    As Integer
    Dim ByteCount       As Long
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    Range = ValidateGetBytes(LocalChars, Index, Count)
    
    ByteCount = GetBytesCore(LocalChars, Range.Index, Range.Count, NullBytes, 0, Nothing)
    SAPtr(GetBytes) = SafeArrayCreateVectorEx(vbByte, 0, ByteCount)
    GetBytesCore LocalChars, Range.Index, Range.Count, GetBytes, 0, Nothing
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    FreeChars LocalChars
    Return
End Function

''
' Encodes a set of characters into an array of bytes, returning the number of bytes produced.
'
' @param Chars the set of characters to be encoded into bytes. This can
' be either an Integer array or a String.
' @param CharIndex The index of the first character to begin encoding from.
' @param CharCount The number of characters to be encoded.
' @param Bytes The destination for the encoded characters.
' @param ByteIndex The index of the first byte stored in the destination array.
' @return The number of bytes produce by the encoding.
' @remarks The <i>Bytes</i> array cannot be null and must be large enough to hold the encoded
' set of characters. To obtain the size required, use GetByteCount to find the necessary size.
' Or, use GetMaxByteCount to get the largest size that could be produced. The largest size is not
' the perfect size, only the maximum number of bytes that could ever be produced by a number of
' characters. Resizing of the resultant <i>Bytes</i> array may be necessary.
'
Public Function GetBytesEx(ByRef Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Dim LocalChars()    As Integer
    Dim Range           As ListRange
    Dim ByteCount       As Long
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    ValidateGetBytesEx LocalChars, CharIndex, CharCount, Bytes, ByteIndex
    
    ByteCount = GetBytesCore(LocalChars, CharIndex, CharCount, NullBytes, 0, Nothing)
    GetBytesEx = GetBytesCore(LocalChars, CharIndex, CharCount, Bytes, ByteIndex, Nothing)
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars LocalChars
    Return
End Function

''
' Returns the number of characters that would be produced by decoding a byte array.
'
' @param Bytes The bytes to use when calculating the number of characters.
' @param Index The starting index in <i>Bytes</i> to being calculating from.
' @param Count The number of bytes to be used in the calculation.
' @return The number of characters that would be decoded from the set of bytes.
'
Public Function GetCharCount(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim Range As ListRange
    
    Range = ValidateGetChars(Bytes, Index, Count)
    
    GetCharCount = GetCharsCore(Bytes, Range.Index, Range.Count, NullChars, 0, Nothing)
End Function

''
' Decodes a set of bytes into a set of characters.
'
' @param Bytes The set of bytes to be decoded into characters.
' @param Index The index of the first byte to be used in decoding.
' @param Count The number of bytes to be used in decoding.
' @return An array of characters decoded from the set of bytes.
'
Public Function GetChars(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Integer()
    Dim Range       As ListRange
    Dim CharCount   As Long
    
    Range = ValidateGetChars(Bytes, Index, Count)
        
    CharCount = GetCharsCore(Bytes, Range.Index, Range.Count, NullChars, 0, Nothing)
    SAPtr(GetChars) = SafeArrayCreateVectorEx(vbInteger, 0, CharCount)
    GetCharsCore Bytes, Range.Index, Range.Count, GetChars, 0, Nothing
End Function

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
Public Function GetCharsEx(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    ValidateGetCharsEx Bytes, ByteIndex, ByteCount, Chars, CharIndex
    GetCharsEx = GetCharsCore(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Nothing)
End Function

''
' Decodes a set of bytes into a String.
'
' @param Bytes The set of bytes to be decoded into a string.
' @param Index The index of the first byte to be decoded.
' @param Count The number of bytes to be used in the decoding.
' @return A string containing the decoded set of bytes.
'
Public Function GetString(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As String
    Dim Range       As ListRange
    Dim CharCount   As Long
    Dim Chars()     As Integer
    
    Range = ValidateGetChars(Bytes, Index, Count)
    
    CharCount = GetCharsCore(Bytes, Range.Index, Range.Count, NullChars, 0, Nothing)
    GetString = SysAllocStringLen(0, CharCount)
    Chars = AllocChars(GetString)
    
    On Error GoTo Catch
    GetCharsCore Bytes, Range.Index, Range.Count, Chars, 0, Nothing
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    FreeChars Chars
    Return
End Function

''
' Returns an array of bytes that represents this encoding.
'
' @return A byte array containg the bytes to be used to identify this encoding type.
' @remarks <p>If the encoding was created setting the <i>ShouldEmitUTF8Identifier</i> paramter
' to true, then this will return the UTF-8 identifer &HFFEF already encoded into the 3
' bytes. The bytes are &HEF, &HBB, &HBF. They will decode to &HFFEF. Since &HFF and &HFE
' are not valid in a UTF-8 encoding, these will unique and can be used to identify the
' encoding used in the stream.</p>
' If <i>ShouldEmitUTF8Identifier</i> was set to False, then an empty array will be returned.
' The array will not be null. It will have a UBound of -1.
'
Public Function GetPreamble() As Byte()
    If mShouldEmitUTF8Identifier Then
        GetPreamble = Cor.NewBytes(&HEF, &HBB, &HBF)
    Else
        GetPreamble = Cor.NewBytes()     ' creates a 0-length array instead of a null array.
    End If
End Function

''
' Returns the maximum number of bytes that can be created from a specific number of characters.
'
' @param CharCount The number of characters to be encoded.
' @return The maximum number of bytes that can be generated from the specified number of characters.
'
Public Function GetMaxByteCount(ByVal CharCount As Long) As Long
    If CharCount < 0 Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_NeedNonNegNum
    
    ' chars are a max 3 bytes each.
    GetMaxByteCount = (CharCount + 1) * 3 ' plus one incase the first char part of a surrogate pair from a previous operation.
End Function

''
' Returns the maximum number of characters than can be decoded from the number of bytes specified.
'
' @param ByteCount The number of bytes to be decoded.
' @return The maximum number of characters that can be decoded from the specified number of bytes.
'
Public Function GetMaxCharCount(ByVal ByteCount As Long) As Long
    If ByteCount < 0 Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_NeedNonNegNum
        
    GetMaxCharCount = ByteCount + 1 ' plus one incase the first byte is the last byte of a surrogate pair from a previous operation.
End Function

''
' Returns an encoder that maintains state.
'
' @return The stateful encoder.
' @remarks Unlike UTF8Encoding, the encoder maintains state between encoding calls. This allows
' for a set of characters to be split up but still be encoded as a single set of characters. This
' will only be effective when encoding UTF-16 surrogate pairs of characters. Otherwise, a single
' characters is all that is required to be encodable. the UTF-16 surrogate pairs can be split to
' the last character of one call and the first character of the next and the encoding will
' continue as if a single call was made.
'
Public Function GetEncoder() As Encoder
    Dim Encoder As New UTF8Encoder
    Encoder.Init Me
    Set GetEncoder = Encoder
End Function

''
' Returns a decoder that maintains state.
'
' @return The stateful decoder.
' @remarks Unlike UTF8Encoding, the decoder maintains state between decoding calls. this allows
' for a sequence of bytes to not have to be in the same call in order to be decoded to the
' representive character. Other than the standard ASCII set of characters (0-127), all other
' characters are represented by 2-4 bytes. Using this decoder allows for some of those bytes to be
' at the end of one call and the rest be at the beginning of the next call. The decoder will
' continue decoding as if all the required bytes were in the same call.
'
Public Function GetDecoder() As Decoder
    Dim Decoder As New UTF8Decoder
    Decoder.Init Me
    Set GetDecoder = Decoder
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        If Value Is Nothing Then Exit Function
        
        If TypeOf Value Is UTF8Encoding Then
            Dim Other As UTF8Encoding
            Set Other = Value
            Equals = (mShouldEmitUTF8Identifier = Other.EmitUTF8Identifier)
        End If
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = CODE_PAGE Or CLng(IIf(mShouldEmitUTF8Identifier, &H80000000, 0)) Or CLng(IIf(mThrowException, &H40000000, 0))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal ShouldEmitUTF8Identifier As Long, ByVal ThrowOnInvalidBytes As Boolean)
    mShouldEmitUTF8Identifier = ShouldEmitUTF8Identifier
    mThrowException = ThrowOnInvalidBytes
    
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

Friend Function GetCharsCore(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal Decoder As UTF8Decoder) As Long
    Dim MaxIndex        As Long
    Dim StartCharIndex  As Long
    Dim MaxCharIndex    As Long
    Dim EncLength       As Long
    Dim BytesRemaining  As Long
    Dim b               As Byte
    Dim Char            As Long
    Dim PutChars        As Boolean
    Dim Fallback        As DecoderFallbackBuffer
    
    If SAPtr(Chars) <> vbNullPtr Then
        PutChars = True
        MaxCharIndex = UBound(Chars)
    End If
    
    StartCharIndex = CharIndex
    MaxIndex = ByteIndex + ByteCount
    
    If Not Decoder Is Nothing Then
        Char = Decoder.Char
        BytesRemaining = Decoder.BytesLeft
        EncLength = Decoder.BytesInSequence
    End If
    
    Do While ByteIndex < MaxIndex
        b = Bytes(ByteIndex)
        
        If EncLength = 0 Then
            If b < &H80 Then
                ' ASCII
                If PutChars Then
                    If CharIndex > MaxCharIndex Then _
                        ThrowDecoderSmallConversionBuffer
                    
                    Chars(CharIndex) = b
                End If
                
                CharIndex = CharIndex + 1
                BytesRemaining = 0
            ElseIf (b And &HE0) = &HC0 Then
                ' 2-byte UTF8 encoding
                Char = b And &H1F   ' decode 5 bits
                
                If Char > 1 Then ' 2 is the minimum the first of the 2 bytes can be.
                    EncLength = 2
                    BytesRemaining = 1
                Else
                    ' trying to create a 2-byte value less than &h80
                    mBytesUnknown(0) = b
                    GoSub InsertFallback
                    BytesRemaining = 0
                End If
            ElseIf (b And &HF0) = &HE0 Then
                ' 3-byte UTF8 encoding
                Char = b And &HF
                EncLength = 3
                BytesRemaining = 2
            ElseIf (b And &HF8) = &HF0 Then
                ' 4-byte (surrogate pair) UTF8 encoding
                Char = b And &H7    ' decode 3 bits
                EncLength = 4
                BytesRemaining = 3
            Else
                ' add fallback length
                mBytesUnknown(0) = b
                GoSub InsertFallback
                BytesRemaining = 0
            End If
        ElseIf (b And &HC0) <> &H80 Then
            ' invalid for all multibyte encodings
            ' add fallback for each byte encoded
            GoSub FallbackDecodedBytes
            
            EncLength = 0
            BytesRemaining = 0
            ByteIndex = ByteIndex - 1
        Else
            Char = (Char * vbShift6Bits) Or (b And &H3F)
            BytesRemaining = BytesRemaining - 1
            
            Select Case EncLength
                Case 2 ' U+0080 - U+07FF
                    If Char > &H7F And Char < &H800 Then
                        If PutChars Then
                            If CharIndex > MaxCharIndex Then _
                                ThrowDecoderSmallConversionBuffer
                                
                            Chars(CharIndex) = Char
                        End If
                        
                        CharIndex = CharIndex + 1
                    Else
                        ByteIndex = ByteIndex - 1
                        mBytesUnknown(0) = Bytes(ByteIndex)
                        GoSub InsertFallback
                    End If
                    
                Case 3 ' U+0800 - U+FFFF
                    If BytesRemaining = 1 Then
                        ' check first 2 bytes conform
                        If Char < &H20 Or IsSurrogate(Char * vbShift6Bits) Then
                            ' first two bytes represent an invalid sequence for 3-byte encoding
                            ' so do fallback for invalid 2-byte sequence and start new sequence.
                            mBytesUnknown(0) = Bytes(ByteIndex - 1)
                            GoSub InsertFallback
                            BytesRemaining = 0
                        End If
                    Else
                        ' no need to validate char because there is no invalid sequence that can make it this far.
                        If PutChars Then
                            If CharIndex > MaxCharIndex Then _
                                ThrowDecoderSmallConversionBuffer
                                
                            Chars(CharIndex) = AsWord(Char)
                        End If
                        
                        CharIndex = CharIndex + 1
                    End If
                    
                Case 4 ' U+D800 - U+DBFF and U+DC00 - U+DFFF encoded as UTF-32 in 21 bits.
                    If BytesRemaining = 0 Then
                        ' we should have decoded both high and low surrogates.
                        Dim HighSurrogate   As Long
                        Dim LowSurrogate    As Long
                        
                        ' convert to two UTF-16 surrogates.
                        LowSurrogate = (Char And &H3FF) + &HDC00&
                        HighSurrogate = (Char \ vbShift10Bits) + &HD7C0&
                        
                        If IsHighSurrogate(HighSurrogate) Then
                            If IsLowSurrogate(LowSurrogate) Then
                                If PutChars Then
                                    If CharIndex + 1 > MaxCharIndex Then _
                                        ThrowDecoderSmallConversionBuffer
                                        
                                    Chars(CharIndex) = AsWord(HighSurrogate)
                                    Chars(CharIndex + 1) = AsWord(LowSurrogate)
                                End If
                                
                                CharIndex = CharIndex + 2
                            Else
                                ' we have an invalid low surrogate and hanging high surrogate so fallback high surrogate
                                ' and restart sequence with invalid low surrogate character bytes.
                                mBytesUnknown(0) = Bytes(ByteIndex - 3)
                                GoSub InsertFallback
                                ByteIndex = ByteIndex - 2
                            End If
                        Else
                            ' we have an invalid high surrogate, so fallback the first byte of the high surrogate
                            ' then rollback index to start sequence with next by after start of original high surrogate sequence.
                            mBytesUnknown(0) = Bytes(ByteIndex - 3)
                            GoSub InsertFallback
                            ByteIndex = ByteIndex - 3
                        End If
                    End If
            End Select
            
            If BytesRemaining = 0 Then
                EncLength = 0
            End If
        End If
        
        ByteIndex = ByteIndex + 1
    Loop
    
    If Decoder Is Nothing Then
        If BytesRemaining > 0 Then
            ' didn't finish decoding character
            GoSub FallbackDecodedBytes
        End If
    ElseIf Decoder.Flushing Then
        If BytesRemaining > 0 Then
            ' didn't finish decoding character
            GoSub FallbackDecodedBytes
        End If
    ElseIf PutChars Then
        ' store state for next call
        Decoder.BytesInSequence = EncLength
        Decoder.Char = Char
        Decoder.BytesLeft = BytesRemaining
    End If
        
    GetCharsCore = CharIndex - StartCharIndex
    Exit Function
    

FallbackDecodedBytes:
    Dim TempIndex As Long
    TempIndex = ByteIndex - (EncLength - BytesRemaining)
    
    Do While TempIndex < ByteIndex
        mBytesUnknown(0) = Bytes(TempIndex)
        GoSub InsertFallback
        TempIndex = TempIndex + 1
    Loop
    
    Return
    
InsertFallback:
    If Fallback Is Nothing Then
        If Decoder Is Nothing Then
            Set Fallback = mDecoderFallback.CreateFallbackBuffer
        Else
            Set Fallback = Decoder.FallbackBuffer
        End If
    End If
    
    If Fallback.Fallback(mBytesUnknown, 0) Then
        Do While Fallback.Remaining > 0
            Dim Ch As Integer
            Ch = Fallback.GetNextChar
            
            If PutChars Then
                If CharIndex > MaxCharIndex Then _
                    ThrowDecoderSmallConversionBuffer
                    
                Chars(CharIndex) = Ch
            End If
            
            CharIndex = CharIndex + 1
        Loop
        
        Fallback.Reset
    End If

    Return
End Function

Friend Function GetBytesCore(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal Encoder As UTF8Encoder) As Long
    Dim Ch          As Long
    Dim PrevChar    As Long
    Dim Fallback    As EncoderFallbackBuffer
    Dim InFallback  As Boolean
    Dim MaxIndex    As Long
    Dim StartIndex  As Long
    Dim MaxByteIndex As Long
    Dim PutBytes    As Boolean
    
    If SAPtr(Bytes) <> vbNullPtr Then
        MaxByteIndex = UBound(Bytes)
        PutBytes = True
    End If
    
    StartIndex = ByteIndex
    MaxIndex = CharIndex + CharCount
    
    If Not Encoder Is Nothing Then
        PrevChar = Encoder.PrevChar
    End If
    
    Do While CharIndex < MaxIndex Or InFallback
        If InFallback Then
            If Fallback.Remaining > 0 Then
                Ch = Fallback.GetNextChar And &HFFFF&
            ElseIf CharIndex < MaxIndex Then
                Ch = Chars(CharIndex) And &HFFFF&
                CharIndex = CharIndex + 1
                InFallback = False
            Else
                Exit Do
            End If
        Else
            Ch = Chars(CharIndex) And &HFFFF&
            CharIndex = CharIndex + 1
        End If
        
        GoSub ProcessChar
    Loop
    
    ' do we have a hanging high surrogate?
    If Encoder Is Nothing Then
        GoSub ProcessPrevChar
    Else
        If Encoder.Flush Then
            GoSub ProcessPrevChar
        End If
    
        If PutBytes Then
            Encoder.PrevChar = PrevChar
        End If
    End If
    
    GetBytesCore = ByteIndex - StartIndex
    Exit Function
    
ProcessPrevChar:
    If PrevChar <> 0 Then
        If FallbackUnknownChar(PrevChar, Fallback, Encoder) Then
            Do While Fallback.Remaining > 0
                Ch = Fallback.GetNextChar And &HFFFF&
                GoSub ProcessChar
            Loop
        End If
        
        PrevChar = 0
    End If

    Return
    
ProcessChar:
    If PrevChar = 0 Then
        Select Case Ch
            Case 0 To &H7F  ' ASCII
                If PutBytes Then
                    If ByteIndex > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer
                    
                    Bytes(ByteIndex) = Ch
                End If
                
                ByteIndex = ByteIndex + 1
            Case &H80 To &H7FF  ' 2-byte UTF8 encoded
                If PutBytes Then
                    If ByteIndex + 1 > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer
                    
                    Bytes(ByteIndex) = &HC0 Or (Ch \ vbShift6Bits)
                    Bytes(ByteIndex + 1) = &H80 Or (Ch And &H3F)
                End If
                
                ByteIndex = ByteIndex + 2
            Case HighSurrogateStart To HighSurrogateEnd
                PrevChar = Ch
            Case LowSurrogateStart To LowSurrogateEnd
                ' a low surrogate without a preceeding high surrogate
                ' just fallback the low surrogate.
                InFallback = FallbackUnknownChar(Ch, Fallback, Encoder)
            Case Else   ' 0800-FFFF excluding surrogates; 3-byte UTF8 encoded
                If PutBytes Then
                    If ByteIndex + 2 > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer
                
                    Bytes(ByteIndex) = &HE0 Or ((Ch \ &H1000&) And &HF)
                    Bytes(ByteIndex + 1) = &H80 Or ((Ch \ vbShift6Bits) And &H3F)
                    Bytes(ByteIndex + 2) = &H80 Or (Ch And &H3F)
                End If
                
                ByteIndex = ByteIndex + 3
        End Select
    Else
        Select Case Ch
            Case LowSurrogateStart To LowSurrogateEnd
                ' we have a surrogate pair
                If PutBytes Then
                    Dim UnicodeChar As Long
                    If ByteIndex + 3 > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer
                    
                    UnicodeChar = ((PrevChar - HighSurrogateStart) * &H400) + (Ch - LowSurrogateStart) + &H10000
                    Bytes(ByteIndex + 3) = &H80 Or (UnicodeChar And &H3F)
                    Bytes(ByteIndex + 2) = &H80 Or ((UnicodeChar \ &H40&) And &H3F)
                    Bytes(ByteIndex + 1) = &H80 Or ((UnicodeChar \ &H1000&) And &H3F)
                    Bytes(ByteIndex) = &HF0 Or (UnicodeChar \ &H40000)
                End If
                
                ByteIndex = ByteIndex + 4
            Case Else
                ' we have a high surrogate with no following low surrogate
                ' so fallback the high surrogate. move the index back so
                ' we will attempt to re-process the current character once
                ' the fallback has completed.
                CharIndex = CharIndex - 1
                InFallback = FallbackUnknownChar(PrevChar, Fallback, Encoder)
        End Select
        
        PrevChar = 0
    End If
    
    Return
End Function

Private Function FallbackUnknownChar(ByVal UnknownChar As Long, ByRef Fallback As EncoderFallbackBuffer, ByVal Encoder As Encoder) As Boolean
    If Fallback Is Nothing Then
        If Encoder Is Nothing Then
            Set Fallback = mEncoderFallback.CreateFallbackBuffer
        Else
            Set Fallback = Encoder.FallbackBuffer
        End If
    End If
    
    FallbackUnknownChar = Fallback.Fallback(UnknownChar)
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
        mThrowException = .ReadProperty(Prop_ThrowException, True)
        mShouldEmitUTF8Identifier = .ReadProperty(Prop_ShouldEmitUTF8Identifier, True)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty Prop_ThrowException, mThrowException
        .WriteProperty Prop_ShouldEmitUTF8Identifier, mShouldEmitUTF8Identifier
    End With
End Sub


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
    
End Property

Private Function Encoding_Clone() As Encoding
    Set Encoding_Clone = Clone
End Function

Private Property Set Encoding_DecoderFallback(ByVal Value As DecoderFallback)

End Property

Private Property Get Encoding_DecoderFallback() As DecoderFallback

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


