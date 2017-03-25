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

Private mThrowException             As Boolean
Private mShouldEmitUTF8Identifier   As Boolean


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
    BodyName = "utf-8"
End Property

''
' Returns the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
'
Public Property Get WebName() As String
    WebName = "utf-8"
End Property

''
' Returns the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
'
Public Property Get HeaderName() As String
    HeaderName = "utf-8"
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
    WindowsCodePage = 1200
End Property

''
' The name of this encoding.
'
' @return The encoding's name.
'
Public Property Get EncodingName() As String
    EncodingName = "Unicode (UTF-8)"
End Property

''
' Indicates if this encoding can be used by browsers to display text.
'
' @return The indication for browser display capable.
'
Public Property Get IsBrowserDisplay() As Boolean
    IsBrowserDisplay = True
End Property

''
' Indicates if this encoding can be used to save data with this encoding.
'
' @return The indication for the browser saving capable.
'
Public Property Get IsBrowserSave() As Boolean
    IsBrowserSave = True
End Property

''
' Indicates if this encoding can be used to display mail and news by
' mail and news clients.
'
' @return Indication for mail and news client capable.
'
Public Property Get IsMailNewsDisplay() As Boolean
    IsMailNewsDisplay = True
End Property

''
' Indicates if this encoding can be used to save data by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
'
Public Property Get IsMailNewsSave() As Boolean
    IsMailNewsSave = True
End Property

''
' Returns the number of bytes that would be produced from the set of characters using this encoding.
'
' @param Chars A set of characters used to calculated the number of bytes once encoded with this encoding.
' @param Index The starting index of the first character to be used.
' @param Count The number of characters to include in the calculation.
' @return The number of bytes that would be produced by encoding the set of characters.
' @remarks The <i>Chars</i> parameter will accept either a String or an Integer array. The Index
' is zero based for both the String and Integer array, not 1 based like Mid$.
'
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim Chrs()  As Integer
    Dim Range   As ListRange
    
    On Error GoTo Catch
    
    Chrs = AsChars(Chars)
    Range = Statics.Encoding.ValidateOptionalCharRange(Chrs, Index, Count)
    
    If Range.Count > 0 Then
        GetByteCount = InternalGetByteCount(Chrs, Range.Index, Range.Count, Nothing)
    End If
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars Chrs
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
    Dim Range       As ListRange
    Dim Chrs()      As Integer
    Dim Output()    As Byte
    
    On Error GoTo Catch
    
    Chrs = AsChars(Chars)
    Range = Statics.Encoding.ValidateOptionalCharRange(Chrs, Index, Count)
    
    If Range.Count > 0 Then
        ReDim Output(0 To InternalGetByteCount(Chrs, Range.Index, Range.Count, Nothing) - 1)
        InternalGetBytes Chrs, Range.Index, Range.Count, Output, 0, Nothing
        GetBytes = Output
    Else
        GetBytes = Cor.NewBytes()
    End If
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars Chrs
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
    Dim Chrs()      As Integer
    Dim Range       As ListRange
    Dim ByteCount   As Long
    
    Chrs = AsChars(Chars)
    
    On Error GoTo Catch
    
    Statics.Encoding.ValidateCharRange Chrs, CharIndex, CharCount
    ByteCount = InternalGetByteCount(Chrs, CharIndex, CharCount, Nothing)
    Statics.Encoding.ValidateByteRange Bytes, ByteIndex, ByteCount
    
    GetBytesEx = InternalGetBytes(Chrs, CharIndex, CharCount, Bytes, ByteIndex, Nothing)
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars Chrs
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
    
    Range = Statics.Encoding.ValidateOptionalByteRange(Bytes, Index, Count)
    
    Dim NullChars() As Integer
    GetCharCount = InternalGetChars(Bytes, Range.Index, Range.Count, NullChars, 0, Nothing)
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
    Dim Output()    As Integer
    
    Range = Statics.Encoding.ValidateOptionalByteRange(Bytes, Index, Count)
    
    If Range.Count > 0 Then
        ReDim Output(0 To GetMaxCharCount(Range.Count))
        
        Dim Size As Long
        Size = InternalGetChars(Bytes, Range.Index, Range.Count, Output, 0, Nothing)
        If Size <> UBound(Output) + 1 Then
            ReDim Preserve Output(0 To Size - 1)
        End If
    Else
        Output = Cor.NewIntegers()
    End If
    
    GetChars = Output
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
    Statics.Encoding.ValidateByteRange Bytes, ByteIndex, ByteCount
    ValidateCharArray Chars
    
    GetCharsEx = InternalGetChars(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Nothing)
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
    Dim Size        As Long
    Dim Result      As String
    Dim Output()    As Integer
    
    Range = Statics.Encoding.ValidateOptionalByteRange(Bytes, Index, Count)
    Size = GetMaxCharCount(Range.Count)
    Result = SysAllocStringLen(0, Size)
    Output = AllocChars(Result)
    
    On Error GoTo Catch
    Size = GetCharsEx(Bytes, Range.Index, Range.Count, Output, 0)
    
    If Len(Result) <> Size Then
        GetString = Left$(Result, Size)
    Else
        GetString = Result
    End If
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars Output
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
Friend Sub Init(ByVal ShouldEmitUTF8Identifier As Long, ByVal ShouldThrowException As Boolean)
    mShouldEmitUTF8Identifier = ShouldEmitUTF8Identifier
    mThrowException = ShouldThrowException
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friends
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get EmitUTF8Identifier() As Boolean
    EmitUTF8Identifier = mShouldEmitUTF8Identifier
End Property

Friend Function InternalGetBytes(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByRef Encoder As UTF8Encoder) As Long
    Const HIGH_OFFSET   As Long = &HD800&
    Const LOW_OFFSET    As Long = &HDC00&
    
    Dim Ch                 As Long
    Dim MaxIndex            As Long
    Dim MaxByteIndex        As Long
    Dim StartIndex          As Long
    Dim HighSurrogateChar   As Long
    
    If CorArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException("Bytes", Environment.GetResourceString(ArgumentNull_Array))
    If ByteIndex < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException("ByteIndex", Environment.GetResourceString(ArgumentOutOfRange_LBound))
    
    If Not Encoder Is Nothing Then
        HighSurrogateChar = Encoder.HighSurrogateChar
    End If
    
    StartIndex = ByteIndex
    MaxByteIndex = UBound(Bytes) + 1
    MaxIndex = CharIndex + CharCount
    
    Do While CharIndex < MaxIndex
        Ch = Chars(CharIndex) And &HFFFF&
        
        If HighSurrogateChar <> 0 Then
            Select Case Ch
                Case &HDC00& To &HDFFF&
                    ' we have a surrogate pair.
                    Dim pair As Long
                    If ByteIndex + 4 > MaxByteIndex Then _
                        SmallBufferError "Bytes"
                    
                    pair = ((HighSurrogateChar - HIGH_OFFSET) * &H400) + (Ch - LOW_OFFSET) + &H10000
                    Bytes(ByteIndex + 3) = &H80 Or (pair And &H3F)
                    Bytes(ByteIndex + 2) = &H80 Or ((pair \ &H40&) And &H3F)
                    Bytes(ByteIndex + 1) = &H80 Or ((pair \ &H1000&) And &H3F)
                    Bytes(ByteIndex) = &HF0 Or (pair \ &H40000)
                    ByteIndex = ByteIndex + 4
                    HighSurrogateChar = 0
                Case &HD800& To &HDBFF&
                    ' we have two high surrogates in a row
                    If mThrowException Then _
                        Throw Cor.NewArgumentException("A high surrogate was followed by a second high surrogate.", "Chars")
                    If ByteIndex + 3 > MaxByteIndex Then _
                        SmallBufferError "Bytes"
                    
                    ByteIndex = EncodeSurrogate(HighSurrogateChar, Bytes, ByteIndex)
                    HighSurrogateChar = Ch
                Case Else
                    If mThrowException Then _
                        Throw Cor.NewArgumentException("A high surrogate was not followed by a low surrogate.", "Chars")
                    If ByteIndex + 3 >= MaxByteIndex Then _
                        SmallBufferError "Bytes"
                    
                    ByteIndex = EncodeSurrogate(HighSurrogateChar, Bytes, ByteIndex)
                    CharIndex = CharIndex - 1
                    HighSurrogateChar = 0
            End Select
        Else
            Select Case Ch
                Case Is < &H80     ' ascii
                    If ByteIndex + 1 > MaxByteIndex Then _
                        SmallBufferError "Bytes"
                    
                    Bytes(ByteIndex) = Ch
                    ByteIndex = ByteIndex + 1
                Case Is < &H800
                    If ByteIndex + 2 > MaxByteIndex Then _
                        SmallBufferError "Bytes"
                    
                    Bytes(ByteIndex) = &HC0 Or (Ch \ 64)
                    Bytes(ByteIndex + 1) = &H80 Or (Ch And &H3F)
                    ByteIndex = ByteIndex + 2
                Case &HD800& To &HDBFF&     ' high surrogate
                    HighSurrogateChar = Ch
                Case &HDC00& To &HDFFF&     ' low surrogate
                    If mThrowException Then _
                        Throw Cor.NewArgumentException("A low surrogate was not preceeded by a high surrogate.", "Chars")
                    If ByteIndex + 3 > MaxByteIndex Then _
                        SmallBufferError "Bytes"
                        
                    ByteIndex = EncodeSurrogate(Ch, Bytes, ByteIndex)
                Case Else
                    If ByteIndex + 3 > MaxByteIndex Then _
                        SmallBufferError "Bytes"
                    
                    ByteIndex = EncodeSurrogate(Ch, Bytes, ByteIndex)
            End Select
        End If
        
        CharIndex = CharIndex + 1
    Loop
    
    If Not Encoder Is Nothing Then
        If Not Encoder.Flush Then
            Encoder.HighSurrogateChar = HighSurrogateChar
        Else
            If mThrowException And (HighSurrogateChar <> 0) Then _
                Throw Cor.NewArgumentException("A high surrogate was not followed by a low surrogate.", "Chars")
            
            Encoder.HighSurrogateChar = 0
        End If
    ElseIf HighSurrogateChar <> 0 Then
        If mThrowException Then _
            Throw Cor.NewArgumentException("A high surrogate was not followed by a low surrogate.", "Chars")
        
        ByteIndex = EncodeSurrogate(HighSurrogateChar, Bytes, ByteIndex)
    End If
    
    InternalGetBytes = ByteIndex - StartIndex
End Function

Friend Function InternalGetChars(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByRef Decoder As UTF8Decoder) As Long
    Dim BytesInSequence As Long
    Dim b               As Byte
    Dim Ch              As Long
    Dim CharCount       As Long
    Dim BytesLeft       As Long
    Dim PutChars        As Boolean
    Dim MaxIndex        As Long
    Dim maxCharIndex    As Long
    
    PutChars = Not CorArray.IsNull(Chars)
    If PutChars Then
        maxCharIndex = UBound(Chars)
        If CharIndex < LBound(Chars) Then _
            Throw Cor.NewArgumentOutOfRangeException("CharIndex", Environment.GetResourceString(ArgumentOutOfRange_LBound))
        If CharIndex > maxCharIndex Then _
            Throw Cor.NewArgumentOutOfRangeException("CharIndex", Environment.GetResourceString(ArgumentOutOfRange_UBound))
    End If
    If Not Decoder Is Nothing Then
        BytesLeft = Decoder.BytesLeft
        Ch = Decoder.Char
        BytesInSequence = Decoder.BytesInSequence
    End If
    
    MaxIndex = ByteIndex + ByteCount

    Do While ByteIndex < MaxIndex
        b = Bytes(ByteIndex)
        
        ' only ascii an character will be 1 byte.   0xxxxxxx format
        If b < &H80 Then
            ' we were expecting more bytes to complete the sequence.
            If BytesInSequence > 0 Then _
                ByteStreamError ByteIndex
                
            If PutChars Then
                If CharIndex > maxCharIndex Then _
                    SmallBufferError "Chars"
                
                Chars(CharIndex) = b
                CharIndex = CharIndex + 1
            End If
            BytesLeft = 0
            BytesInSequence = 0
            CharCount = CharCount + 1
        ElseIf BytesInSequence = 0 Then
            Ch = b
            ' we count off the number of bits set starting from
            ' the highest bit. No more than 4 should be found.
            ' This will tell us how many bytes need to be put together
            ' to form the output character, or surrogate pair.
            Do While (b And &H80)
                BytesInSequence = BytesInSequence + 1
                b = (b And &H7F) * 2
            Loop
            BytesLeft = BytesInSequence - 1
            Select Case BytesInSequence
                Case 2      ' 110xxxxx format
                    Ch = Ch And &H1F
                    If Ch < 2 Then _
                        ByteStreamError ByteIndex
                Case 3      ' 1110xxxx format
                    Ch = Ch And &HF
                Case 4      ' 11110xxx format
                    Ch = Ch And &H7
                    If Ch <> 0 Then
                        If Ch <> &H4 Then
                            ByteStreamError ByteIndex
                        End If
                    End If
                Case Else
                    ByteStreamError ByteIndex
                    
                    If PutChars Then
                        If CharIndex > maxCharIndex Then
                            SmallBufferError "Chars"
                        End If
                        
                        Chars(CharIndex) = &HFFFD
                        CharIndex = CharIndex + 1
                    End If
                    
                    BytesInSequence = 0
                    CharCount = CharCount + 1
            End Select
        Else
            ' if we get in here then we must be on bytes 2-4 in a byte sequence,
            ' so they all will look like 10xxxxxx.
            If (b And &HC0) <> &H80 Then _
                ByteStreamError ByteIndex
            
            Ch = Ch * &H40
            Ch = Ch + (b And &H3F)
            BytesLeft = BytesLeft - 1
            
            Select Case BytesInSequence
                Case 2      ' 110xxxxx 10xxxxxx -  char from &h80 to &h7ff
                    Select Case Ch
                        Case &H80 To &H7FF
                            If BytesLeft = 0 Then
                                If PutChars Then
                                    If CharIndex > maxCharIndex Then
                                        SmallBufferError "Chars"
                                    End If
                                    
                                    Chars(CharIndex) = Ch
                                    CharIndex = CharIndex + 1
                                End If
                                
                                BytesInSequence = 0
                                CharCount = CharCount + 1
                            End If
                            
                        Case Else
                            ByteStreamError ByteIndex
                            If PutChars Then
                                If CharIndex > maxCharIndex Then
                                    SmallBufferError "Chars"
                                End If
                                
                                Chars(CharIndex) = &HFFFD
                                Chars(CharIndex + 1) = &HFFFD
                                CharIndex = CharIndex + 2
                            End If
                            
                            BytesInSequence = 0
                            CharCount = CharCount + 2
                    End Select
                Case 3      ' 1110xxxx 10xxxxxx 10xxxxxx - char from &h800 to &hffff. We throw on &hd800-&hdfff for bad surrogates
                    If Ch = 0 Then _
                        ByteStreamError ByteIndex
                    
                    Select Case BytesLeft
                        Case 0
                            If PutChars Then
                                If CharIndex > maxCharIndex Then _
                                    SmallBufferError "Chars"
                                
                                Chars(CharIndex) = AsWord(Ch)
                                CharIndex = CharIndex + 1
                            End If
                            BytesInSequence = 0
                            CharCount = CharCount + 1
                        Case 1  ' check for shortest form conformity
                            If Ch < &H20 Or Char.IsSurrogate(Ch * &H40) Then
                                ByteStreamError ByteIndex
                            End If
                            
                    End Select
                Case 4      ' 11110XXX 10XXxxxx 10xxxxxx 10xxxxxx - used for surrogate pairs
                    If BytesLeft = 0 Then
                        Dim hs As Long
                        Dim ls As Long
                        
                        ls = (Ch And &H3FF) + &HDC00&
                        hs = (Ch \ &H400) + &HD7C0&
                        If ls < &HDC00& Or ls > &HDFFF& Then _
                            ByteStreamError ByteIndex
                        If hs < &HD800& Or hs > &HDBFF& Then _
                            ByteStreamError ByteIndex
                        
                        If PutChars Then
                            If CharIndex + 1 > maxCharIndex Then _
                                SmallBufferError "Chars"
                            
                            Chars(CharIndex + 1) = AsWord(ls)   ' AsWord is used to shove the highest bit in
                            Chars(CharIndex) = AsWord(hs)       ' without causing an overflow. The char will
                            CharIndex = CharIndex + 2           ' be negative afterwards.
                        End If
                        
                        CharCount = CharCount + 2
                        BytesInSequence = 0
                    Else
                        ' check for shortest form conformity
                        If Ch = 0 Then _
                            ByteStreamError ByteIndex
                    End If
                Case Else
                    ByteStreamError ByteIndex
                    BytesInSequence = 0
                    BytesLeft = 0
            End Select
        End If
        ByteIndex = ByteIndex + 1
    Loop
        
    If Not Decoder Is Nothing Then
        Decoder.BytesLeft = BytesLeft
        If BytesLeft = 0 Then
            Decoder.Char = 0
        Else
            Decoder.Char = Ch
        End If
        Decoder.BytesInSequence = BytesInSequence
    End If
    InternalGetChars = CharCount
End Function

Friend Function InternalGetByteCount(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByRef Encoder As UTF8Encoder) As Long
    Dim Ret As Long
    Dim HasHighSurrogate As Boolean
    
    If Index < LBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Environment.GetResourceString(ArgumentOutOfRange_LBound))
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If Index + Count - 1 > UBound(Chars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen), "Count")

    If Not Encoder Is Nothing Then HasHighSurrogate = (Encoder.HighSurrogateChar <> 0)
    
    Do While Count > 0
        Select Case Chars(Index)
            Case 0 To &H7F      ' ascii value
                If HasHighSurrogate Then
                    If mThrowException Then _
                        Throw Cor.NewArgumentException("An incorrect low surrogate was found immediately after a high surrogate.", "Chars")
                End If
                Ret = Ret + 1
            Case &H80 To &H7FF     ' 2 byte UTF-8 encoding
                If HasHighSurrogate Then
                    If mThrowException Then _
                        Throw Cor.NewArgumentException("An incorrect low surrogate was found immediately after a high surrogate.", "Chars")
                End If
                Ret = Ret + 2
            Case &HD800 To &HDBFF    ' high surrogate value
                If HasHighSurrogate Then
                    If mThrowException Then _
                        Throw Cor.NewArgumentException("Second high surrogate in a row was found.", "Chars")
                Else
                    HasHighSurrogate = True
                End If
                Ret = Ret + 3
            Case &HDC00 To &HDFFF   ' low surrogate value
                If HasHighSurrogate Then
                    Ret = Ret + 1
                    HasHighSurrogate = False
                Else
                    If mThrowException Then _
                        Throw Cor.NewArgumentException("Low surrogate was found without a precedding high surrogate.", "Chars")
                    
                    Ret = Ret + 3
                End If
            Case Else
                Ret = Ret + 3
        End Select
        Index = Index + 1
        Count = Count - 1
    Loop
    
    If HasHighSurrogate Then
        If Not Encoder Is Nothing Then
            If Encoder.Flush Then
                If mThrowException Then _
                    Throw Cor.NewArgumentException("A high surrogate character was not followed by a low surrogate character.", "Chars")
            End If
        Else
            If mThrowException Then _
                Throw Cor.NewArgumentException("Missing Low Surrogage.", "Chars")
        End If
    End If
    
    InternalGetByteCount = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SmallBufferError(ByRef ParamName As String)
    Error.ArgumentOutOfRange ParamName, Argument_SmallConversionBuffer
End Sub

Private Function EncodeSurrogate(ByVal Char As Long, ByRef Bytes() As Byte, ByVal Index As Long) As Long
    Bytes(Index) = &HE0 Or ((Char \ &H1000&) And &HF)
    Bytes(Index + 1) = &H80 Or ((Char \ &H40&) And &H3F)
    Bytes(Index + 2) = &H80 Or (Char And &H3F)
    EncodeSurrogate = Index + 3
End Function

Private Sub ByteStreamError(ByVal Index As Long)
    If mThrowException Then
        Throw Cor.NewArgumentException(CorString.Format("An invalid byte was found at index {0}.", Index), "Bytes")
    End If
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
Private Property Set Encoding_DecoderFallback(ByVal Value As DecoderFallback)

End Property

Private Property Get Encoding_DecoderFallback() As DecoderFallback

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


