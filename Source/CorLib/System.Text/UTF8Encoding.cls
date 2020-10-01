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
Private mLongCharsSA                As SafeArray1d
Private mIntegerBytesSA             As SafeArray1d
Private mLongCharsSAPtr             As Long
Private mIntegerBytesSAPtr          As Long
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
' Gets the current DecoderFallback instance used by the encoding.
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
' Gets the current EncoderFallback instance used by the encoding.
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
' Calculates the number of bytes produced by encoding the characters in the specified String or Integer().
'
' @param Chars The Integer() or String containing the set of characters to encode.
' @param Index The index of the first character to encode.
' @param Count The number of characters to encode.
' @return The number of bytes produced by encoding the specified of characters.
' @include "..\DocIncludes\UTF8Encoding.GetByteCount.html"
' @see #GetBytes
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
    
    ByteCount = GetByteCountCore(LocalChars, Range.Index, Range.Count, Nothing)
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
    
    GetCharCount = GetCharCountCore(Bytes, Range.Index, Range.Count, Nothing)
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
        
    CharCount = GetCharCountCore(Bytes, Range.Index, Range.Count, Nothing)
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
    
    CharCount = GetCharCountCore(Bytes, Range.Index, Range.Count, Nothing)
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
    Dim Other As UTF8Encoding
    
    On Error GoTo Catch
    Set Other = Value
    Equals = Other.EmitUTF8Identifier = mShouldEmitUTF8Identifier And _
        Other.DecoderFallback.Equals(mDecoderFallback) And _
        Other.EncoderFallback.Equals(mEncoderFallback)
        
Catch:
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
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
    mLongCharsSA.cbElements = 4
    mLongCharsSA.cDims = 1
    mLongCharsSA.cElements = &H7FFFFFFF
    mLongCharsSAPtr = VarPtr(mLongCharsSA)
    
    mIntegerBytesSA.cbElements = 2
    mIntegerBytesSA.cDims = 1
    mIntegerBytesSA.cElements = &H7FFFFFFF
    mIntegerBytesSAPtr = VarPtr(mIntegerBytesSA)
    
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
    
    MaxIndex = Index + Count - 1
    
    If Not Decoder Is Nothing Then
        Char = Decoder.Char
    End If
    
    Do While Index <= MaxIndex
        b = Bytes(Index)
        Index = Index + 1

        If Char = 0 Then
            ' Check if we have an ASCII character we can just stuff immediately.
            If b < &H80 Then
                CharCount = CharCount + 1
                
            ' The start of all multi-byte sequence has the 6-bit set. So check
            ' that here for a valid multi-byte sequence.
            ElseIf (b And &HE0) = &HC0 Then
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
        ElseIf (b And &HC0) = &H80 Then
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
    Dim FastIndex       As Long
    Dim MaxFastIndex    As Long
    Dim CanFastLoop     As Boolean
    Dim BytesAddress    As Long
        
    StartCharIndex = CharIndex
    MaxByteIndex = ByteIndex + ByteCount - 1
    MaxCharIndex = UBound(Chars)
    CanFastLoop = Len1D(Chars) > 1 And Len1D(Bytes) > 3
    
    If CanFastLoop Then
        BytesAddress = VarPtr(Bytes(ByteIndex)) - LBound(Bytes)
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
                
            ' The start of all multi-byte sequence has the 6-bit set. So check
            ' that here for a valid multi-byte sequence.
            ElseIf (b And &HE0) = &HC0 Then
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
        ElseIf (b And &HC0) = &H80 Then
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
    Dim LongChars()     As Long
    Dim CanFastLoop     As Boolean
    Dim CharsAddress    As Long
    Dim ByteCount       As Long
    
    MaxCharIndex = Index + Count - 1
    CanFastLoop = Len1D(Chars) > 0
    
    If CanFastLoop Then
        CharsAddress = VarPtr(Chars(Index)) - LBound(Chars) * 2
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
            
            ' Quick check to see if we should even attempt a fast loop because
            ' we won't know for sure until everything has been set up then
            ' it needs to be torn down. This causes huge slow-downs for non-ascii characters.
            If Ch < &H80 Then
                If Not InFallback And PrevChar = 0 And CanFastLoop Then
                    ' do fast loop
                    If ((CharsAddress + Index) And &H3) = 0 Then ' 32-bit aligned
                        Dim MaxFastIndex As Long
                        Dim FastIndex As Long
                        
                        MaxFastIndex = (Count - Index) \ 2 - 1
                        mLongCharsSA.PVData = CharsAddress + Index
                    
                        SAPtr(LongChars) = mLongCharsSAPtr
                        
                        For FastIndex = 0 To MaxFastIndex
                            Dim Temp As Long
                            
                            Temp = LongChars(FastIndex)
                            
                            If (Temp And &HFF80FF80) <> 0 Then
                                Exit For
                            End If
                        Next
                        
                        SAPtr(LongChars) = vbNullPtr
                        
                        If FastIndex > 0 Then
                            Index = Index + FastIndex * 2
                            ByteCount = ByteCount + FastIndex * 2
                            GoTo ProcessNext
                        End If
                    End If
                End If
            End If
            
            Index = Index + 1
        End If
        
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
    Dim LongChars()     As Long
    Dim IntegerBytes()  As Integer
    Dim CanFastLoop     As Boolean
    Dim CharsAddress    As Long
    Dim BytesAddress    As Long

    MaxByteIndex = UBound(Bytes)
    StartByteIndex = ByteIndex
    MaxCharIndex = CharIndex + CharCount - 1
    CanFastLoop = Len1D(Chars) > 0 And Len1D(Bytes) > 1
    
    If CanFastLoop Then
        CharsAddress = VarPtr(Chars(CharIndex)) - LBound(Chars) * 2
        BytesAddress = VarPtr(Bytes(ByteIndex)) - LBound(Bytes)
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
            
            ' Quick check to see if we should even attempt a fast loop because
            ' we won't know for sure until everything has been set up then
            ' it needs to be torn down. This causes huge slow-downs for non-ascii characters.
            If Ch < &H80 Then
                If Not InFallback And PrevChar = 0 And CanFastLoop Then
                    ' do fast loop
                    If ((BytesAddress + ByteIndex) And &H1) = 0 Then ' 16-bit aligned
                        If ((CharsAddress + CharIndex) And &H3) = 0 Then ' 32-bit aligned
                            Dim MaxFastIndex As Long
                            Dim FastIndex As Long
                            
                            If (MaxByteIndex - ByteIndex) < (MaxCharIndex - CharIndex) Then
                                MaxFastIndex = (MaxByteIndex - ByteIndex) \ 2 - 1
                            Else
                                MaxFastIndex = (CharCount - CharIndex) \ 2 - 1
                            End If
                            
                            mLongCharsSA.PVData = CharsAddress + CharIndex
                            mIntegerBytesSA.PVData = BytesAddress + ByteIndex
                        
                            SAPtr(LongChars) = mLongCharsSAPtr
                            SAPtr(IntegerBytes) = mIntegerBytesSAPtr
                            
                            For FastIndex = 0 To MaxFastIndex
                                Dim Temp As Long
                                
                                Temp = LongChars(FastIndex)
                                
                                ' Check for adjacent ascii characters. If we don't
                                ' have that then we can't store anything, so just
                                ' bail and let the slow loop take care of things.
                                If (Temp And &HFF80FF80) <> 0 Then
                                    Exit For
                                End If

                                #If Release Then
                                IntegerBytes(FastIndex) = ((Temp And &H7F0000) \ vbShift8Bits) Or (Temp And &H7F)
                                #Else
                                Dim Temp2 As Long
                                Temp2 = ((Temp And &H7F0000) \ vbShift8Bits) Or (Temp And &H7F)

                                IntegerBytes(FastIndex) = AsWord(Temp2)
                                #End If
                            Next
                            
                            SAPtr(LongChars) = vbNullPtr
                            SAPtr(IntegerBytes) = vbNullPtr
                            
                            ' If the fast loop successfully processed as least one
                            ' set of adjacent ascii characters, then let the slow
                            ' loop start again with the non-ascii character found.
                            If FastIndex > 0 Then
                                CharIndex = CharIndex + FastIndex * 2
                                ByteIndex = ByteIndex + FastIndex * 2
                                GoTo Continue
                            End If
                        End If
                    End If
                End If
            End If
            
            CharIndex = CharIndex + 1
        End If
        
        If PrevChar = 0 Then
            Select Case Ch
                Case Is < &H80
                    If ByteIndex > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer
                    
                    Bytes(ByteIndex) = Ch
                    ByteIndex = ByteIndex + 1
                Case Is < &H800
                    If ByteIndex > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer

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

