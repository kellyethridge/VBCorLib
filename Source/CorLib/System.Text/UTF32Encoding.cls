VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF32Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2020 Kelly Ethridge
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
' Module: UTF32Encoding
'

''
' Represents a UTF-32 encoding of Unicode characters.
'
' @see ASCIIEncoding
' @see UTF7Encoding
' @see UTF8Encoding
' @see UnicodeEncoding
' @include "..\DocIncludes\UTF32Encoding.html"
Option Explicit
Implements IObject
Implements Encoding

Private Const IsReadOnlyProp        As String = "IsReadOnly"
Private Const EncoderFallbackProp   As String = "EncoderFallback"
Private Const DecoderFallbackProp   As String = "DecoderFallback"

Private mIsReadOnly         As Boolean
Private mIsBigEndian        As Boolean
Private mEmitBOM            As Boolean
Private mCodePage           As Long
Private mEncoderFallback    As EncoderFallback
Private mDecoderFallback    As DecoderFallback


''
' Gets a value indicating whether the current encoding is read-only.
'
' @return <b>True</b> if the current <a href="Encoding.html">Encoding</a> is read-only; otherwise, <b>False</b>.
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mIsReadOnly
End Property

''
' Creates a clone of the current Encoding instance.
'
' @return A cloned version of this instance.
' @remarks A cloned instance will have the read-only limitations removed.
Public Function Clone() As UTF32Encoding
    Set Clone = New UTF32Encoding
    Clone.InitClone Me, mIsBigEndian, mEmitBOM
End Function

''
' Gets the current <a href="EncoderFallback.html">EncoderFallback</a> instance used by the encoding.
'
' @return An object derived from <a href="EncoderFallback.html">EncoderFallback</a>.
' @remarks The default will return an instance of <a href="EncoderReplacementFallback.html">EncoderReplacementFallback</a>.
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
' Gets the current <a href="DecoderFallback.html">DecoderFallback</a> instance used by the encoding.
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
' Gets if the current encoding uses single-byte code points.
'
' @return <b>True</b> for single-byte, <b>False</> for multi-byte code points.
' @remarks This always returns <b>False</b>.
Public Property Get IsSingleByte() As Boolean
    IsSingleByte = False
End Property

''
' Gets the name of the current encoding that can be used with the mail agent body tags.
'
' @return A name for the current <a href="Encoding.html">Encoding</a> that can be used with mail agent body tags.<br/>
' -or-<br/>
' An empty string (""), if the current <a href="Encoding.html">Encoding</a> cannot be used.
' @see Encoding
Public Property Get BodyName() As String
    BodyName = Table.BodyName(CodePage)
End Property

''
' gets the code page identifier of the current <a href="Encoding.html">Encoding</a>.
'
' @return Code page identifier of the current <a href="Encoding.html">Encoding</a>
' @see Encoding
Public Property Get CodePage() As Long
    CodePage = mCodePage
End Property

''
' gets the human-readable description of the current encoding.
'
' @return The human-readable description of the current <a href="Encoding.html">Encoding</a>.
' @remarks The <b>EncodingName</b> property is intended for display. To find a name that can be
' passed to the <a href="EncodingStatic.GetEncoding.html">GetEncoding</a> method, use the <a href="Encoding.Get.WebName.html">WebName</a> property.
' @see Encoding
Public Property Get EncodingName() As String
    EncodingName = Table.EncodingName(CodePage)
End Property

''
' gets a name for the current encoding that can be used with mail agent header tags.
'
' @return A name for the current <a href="Encoding.html">Encoding</a> that can be used with mail agent body tags.<br/>
' -or-<br/>
' An empty string (""), if the current <a href="Encoding.html">Encoding</a> cannot be used.
' @see Encoding
Public Property Get HeaderName() As String
    HeaderName = Table.HeaderName(CodePage)
End Property

''
' Gets if this encoding can be used by browsers to display text.
'
' @return The indication for browser display capable.
' @see Encoding
Public Property Get IsBrowserDisplay() As Boolean
    IsBrowserDisplay = Table.IsBrowserDisplay(CodePage)
End Property

''
' Gets if this encoding can be used to save data with this encoding.
'
' @return The indication for the browser saving capable.
' @see Encoding
Public Property Get IsBrowserSave() As Boolean
    IsBrowserSave = Table.IsBrowserSave(CodePage)
End Property

''
' Gets if this encoding can be used to display mail and news by
' mail and news clients.
'
' @return Indication for mail and news client capable.
' @see Encoding
Public Property Get IsMailNewsDisplay() As Boolean
    IsMailNewsDisplay = Table.IsMailNewsDisplay(CodePage)
End Property

''
' Gets if this encoding can be used to save date by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
' @see Encoding
Public Property Get IsMailNewsSave() As Boolean
    IsMailNewsSave = Table.IsMailNewsSave(CodePage)
End Property

''
' Gets the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered IANA name for the current <a href="Encoding.html">Encoding</a>.
' @see Encoding
Public Property Get WebName() As String
    WebName = Table.WebName(CodePage)
End Property

''
' Gets the Windows Operating System's code page for this encoding.
'
' @return The Windows code page for this encoding.
' @see Encoding
Public Property Get WindowsCodePage() As Long
    WindowsCodePage = Table.WindowsCodePage(CodePage)
End Property

''
' Returns the number of bytes that would be produced from the set of characters using this encoding.
'
' @param Chars The character array containing the set of characters to encode.
' @param Index The index of the first character to encode.
' @param Count The number of characters to encode.
' @return The number of bytes produced by encoding the specified characters.
' @remarks The <i>Chars</i> parameter will accept either a String or an Integer array.
' <p>If <i>Chars</i> is a string then the <i>Index</i> parameter is zero-based.</p>
' @see Encoding
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim LocalChars()    As Integer
    Dim Range           As ListRange
    
    LocalChars = AsChars(Chars)

    On Error GoTo Catch
    Range = ValidateGetBytes(LocalChars, Index, Count)
    GetByteCount = GetByteCountCore(LocalChars, Range.Index, Range.Count, Nothing)
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
' @param Chars The set of characters to be encoded into an array of bytes.
' @param Index The index of the first character to begin encoding from.
' @param Count The number of characters to be encoded.
' @return A byte array containing the encoded characters as bytes.
' @remarks The <i>Chars</i> parameter will accept either a <b>String</b> or an <b>Integer</b> array.
' <p>If <i>Chars</i> is a string then the <i>Index</i> parameter is zero-based.</p>
' @see Encoding
Public Function GetBytes(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    Dim LocalChars()    As Integer
    Dim Range           As ListRange
    Dim ByteCount       As Long
    
    LocalChars = AsChars(Chars)

    On Error GoTo Catch

    Range = ValidateGetBytes(LocalChars, Index, Count)
    ByteCount = GetByteCountCore(LocalChars, Range.Index, Range.Count, Nothing)
    SAPtr(GetBytes) = SafeArrayCreateVectorEx(vbByte, 0, ByteCount)
    
    If ByteCount > 0 Then
        GetBytesCore LocalChars, Range.Index, Range.Count, GetBytes, 0, Nothing
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
' Encodes a set of characters into an array of bytes, returning the number of bytes produced.
'
' @param Chars the set of characters to be encoded into bytes.
' @param CharIndex The index of the first character to begin encoding from.
' @param CharCount The number of characters to be encoded.
' @param Bytes The destination for the encoded characters.
' @param ByteIndex The index of the first byte stored in the destination array.
' @return The number of bytes produce by the encoding.
' @remarks The <i>Bytes</i> array cannot be null and must be large enough to hold the encoded
' set of characters. To obtain the size required, use <a href="UTF32Encoding.GetByteCount.html">GetByteCount</a> to find the necessary size.
' Or, use <a href="UTF32Encoding.GetMaxByteCount.html">GetMaxByteCount</a> to get the largest size that could be produced. The largest size is not
' the perfect size, only the maximum number of bytes that could ever be produced by a number of
' characters. Resizing of the resultant <i>Bytes</i> array may be necessary.
' <p>The <i>Chars</i> parameter will accept either a <b>String</b> or an <b>Integer</b> array.</p>
' <p>If <i>Chars</i> is a string then the <i>CharIndex</i> parameter is zero-based.</p>
' @see Encoding
Public Function GetBytesEx(ByRef Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Dim LocalChars() As Integer

    LocalChars = AsChars(Chars)

    On Error GoTo Catch
    ValidateGetBytesEx LocalChars, CharIndex, CharCount, Bytes, ByteIndex

    If CharCount > 0 Then
        GetBytesEx = GetBytesCore(LocalChars, CharIndex, CharCount, Bytes, ByteIndex, Nothing)
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
' Returns the number of characters that would be produced by decoding a byte array.
'
' @param Bytes The bytes to use when calculating the number of characters.
' @param Index The starting index in <i>Bytes</i> to being calculating from.
' @param Count The number of bytes to be used in the calculation.
' @return The number of characters that would be decoded from the set of bytes.
' @see Encoding
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
' @see Encoding
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
' be decoded. To ensure the <i>Chars</i> array is large enough, use either <a href="ASCIIEncoding.GetCharCount.html">GetCharCount</a> or
' <a href="ASCII.Encoding.GetMaxCharCount.html">GetMaxCharCount</a> to determine a size that will hold the decoded bytes.
' @see Encoding
Public Function GetCharsEx(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    ValidateGetCharsEx Bytes, ByteIndex, ByteCount, Chars, CharIndex
    GetCharsEx = GetCharsCore(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Nothing)
End Function

''
' Returns the maximum number of characters than can be decoded from the number of bytes specified.
'
' @param ByteCount The number of bytes to be decoded.
' @return The maximum number of characters that can be decoded from the specified number of bytes.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>ByteCount</i> is less than zero.<br/>
'           -or-<br/>
'           The resulting number of bytes is greater than the maximum number that can be returned as an integer.
'        </td>
'   </tr>
' </table>
' @see Encoding
Public Function GetMaxCharCount(ByVal ByteCount As Long) As Long
    Dim CharCount As Currency

    If ByteCount < 0 Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_NeedNonNegNum

    ' UTF-32 it represented by 4 bytes and we can fit a high and low surrogate in those
    ' 4 bytes, so the max number of characters per 4 bytes is 2 characters. We divide by
    ' two to represent this. We also add two more characters for a possible high/low
    ' surrogate pair being partially held in a Decoder waiting to be used in the next
    ' GetChars call using that Decoder. If the cached bytes (1 to 3) are used with the
    ' next bytes (1 to 3) to get 4 bytes, it could represent a high/low surrogate pair.
    CharCount = ByteCount \ 2 + 2

    ' We could have every 4 bytes be an invalid Unicode character causing the fallback
    ' characters to be output, so if the fallback can produce more than 2 characters
    ' per 4 bytes, we use that as our new worst case calculation.
    If mDecoderFallback.MaxCharCount > 2 Then
        ' Calculate max by multiplying max bad characters by the max number
        ' of fallback characters that will replace them.
        CharCount = CharCount * mDecoderFallback.MaxCharCount
        
        ' Multiplying by characters is the same as multiplying by twice the ByteCount
        ' value, because each character is 4 bytes.
        CharCount = CharCount \ 2
    End If

    If CharCount > CCur(&H7FFFFFFF) Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_GetCharCountOverflow

    GetMaxCharCount = CharCount
End Function

''
' Returns the maximum number of bytes that can be created from a specific number of characters.
'
' @param CharCount The number of characters to be encoded.
' @return The maximum number of bytes that can be generated from the specified number of characters.
' @see Encoding
Public Function GetMaxByteCount(ByVal CharCount As Long) As Long
    Dim ByteCount As Currency

    If CharCount < 0 Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_NeedNonNegNum

    ' Add one to account for a possible high surrogate being cached in
    ' an Encoder waiting to be used in the next GetBytes call using that
    ' Encoder. If the next character is a low surrogate then the two will
    ' count as a single character, however, if the cached high surrogate
    ' is left hanging, then it needs to be counted on its own. We do that here.
    ByteCount = CharCount + 1

    If mEncoderFallback.MaxCharCount > 1 Then
        ByteCount = ByteCount * mEncoderFallback.MaxCharCount
    End If
    
    ByteCount = ByteCount * 4

    If ByteCount > CCur(&H7FFFFFFF) Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_GetByteCountOverflow

    GetMaxByteCount = ByteCount
End Function

''
' Obtains an encoder that converts a sequence of Unicode characters into an ASCII encoded sequence of bytes.
'
' @return An <a href="Encoder.html">Encoder</a> that converts a sequence of Unicode characters into an ASCII encoded sequence of bytes.
' @remarks The <a href="Encoder.GetBytes.html">Encoder.GetBytes</a> method converts sequential blocks of characters into sequential blocks of bytes,
' in a manner similar to the <a href="ASCIIEncoding.GetBytes.html">GetBytes</a> method of this class. However, an <a href="Encoder.html">Encoder</a> maintains state information between
' calls so it can correctly encode character sequences that span blocks. The <a href="Encoder.html">Encoder</a> also preserves trailing characters
' at the end of data blocks and uses the trailing characters in the next encoding operation. For example, a data block might
' end with an unmatched high surrogate, and the matching low surrogate might be in the next data block. Therefore, <a href="ASCIIEncoding.GetDecoder.html>GetDecoder</a>
' and <b>GetEncoder</b> are useful for network transmission and file operations, because those operations often deal with blocks of data instead of a complete data stream.
' @see Encoding
Public Function GetEncoder() As Encoder
    Dim Encoder As New UTF32Encoder
    
    Encoder.Init Me
    Set GetEncoder = Encoder
End Function

''
' Obtains a decoder that converts an ASCII encoded sequence of bytes into a sequence of Unicode characters.
'
' @return A <a href="Decoder.html">Decoder</a> that converts an ASCII encoded sequence of bytes into a sequence of Unicode characters.
' @remarks The <a href="Decoder.GetChars.html">Decoder.GetChars</a> method converts sequential blocks of bytes into sequential blocks of characters,
' in a manner similar to the <a href="ASCIIEncoding.GetChars.html">GetChars</a> method of this class. However, a <a href="Decoder.html">Decoder</a> maintains state information between calls
' so it can correctly decode byte sequences that span blocks. The Decoder also preserves trailing bytes at the end of data blocks
' and uses the trailing bytes in the next decoding operation. Therefore, <b>GetDecoder</b> and <a href="ASCIIEncoding.GetEncoder.html">GetEncoder</a> are useful for network transmission
' and file operations, because those operations often deal with blocks of data instead of a complete data stream.
' @see Encoding
Public Function GetDecoder() As Decoder
    Dim Decoder As New UTF32Decoder
    
    Decoder.Init Me
    Set GetDecoder = Decoder
End Function

''
' Returns a Unicode byte order mark encoded in UTF-32 format, if the <a href="UTF32Encoding.html">UTF32Encoding</a> object is configured to supply one.
'
' @return A byte array containing the Unicode byte order mark, if the <a href="UTF32Encoding.html">UTF32Encoding</a> object is configured to supply one. Otherwise, this method returns a zero-length byte array
' <p>See <a href="Encoding.GetPreamble.html">Encoding.GetPreamble</a> for additional information.</p>
' @see Encoding
Public Function GetPreamble() As Byte()
    If mEmitBOM Then
        If mIsBigEndian Then
            GetPreamble = Cor.NewBytes(&H0, &H0, &HFE, &HFF)
        Else
            GetPreamble = Cor.NewBytes(&HFF, &HFE, &H0, &H0)
        End If
    Else
        GetPreamble = Cor.NewBytes()
    End If
End Function

''
' Decodes a range of bytes into a <b>String</b>.
'
' @param Bytes The set of bytes to be decoded into a string.
' @param Index The index of the first byte to be decoded.
' @param Count The number of bytes to be used in the decoding.
' @return A string containing the decoded set of bytes.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Bytes</i> is <b>null</b>.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>Index</i> is less than the lower-bound of <i>Bytes</i>.<br/>
'           -or-<br/>
'           <i>Count</i> is less than zero.<br/>
'           -or-<br/>
'           <i>Index</i> and <i>Count</i> do not denote a valid range in <i>Bytes</i>.
'        </td>
'    </tr>
'    <tr>
'        <td><a href="DecoderFallbackException.html">DecoderFallbackException</a></td>
'        <td>
'           A fallback occurred<br/>
'           -and-<br/>
'           <a href="Encoding.Get.DecoderFallback.html">DecoderFallback</a> is set to <a href="DecoderExceptionFallback.html">DecoderExceptionFallback</a>.
'        </td>
'    </tr>
' </table>
' @see Encoding
Public Function GetString(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As String
    Dim Chars() As Integer
    Dim Length  As Long
    Dim Range   As ListRange

    Range = ValidateGetChars(Bytes, Index, Count)
    
    Length = GetCharCountCore(Bytes, Range.Index, Range.Count, Nothing)
    GetString = SysAllocStringLen(0, Length)
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
' Returns a string representation of this object instance.
'
' @return String representing this instance.
' @see IObject
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Dim Other As UTF32Encoding
    
    On Error GoTo Catch
    Set Other = Value
    Equals = CodePage = Other.CodePage And _
             mEncoderFallback.Equals(Other.EncoderFallback) And _
             mDecoderFallback.Equals(Other.DecoderFallback)
    
Catch:
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
' @see IObject
'
Public Function GetHashCode() As Long
    GetHashCode = CodePage Xor mEncoderFallback.GetHashCode Xor mDecoderFallback.GetHashCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal BigEndian As Boolean, ByVal ByteOrderMark As Boolean, ByVal ThrowOnInvalidBytes As Boolean)
    mEmitBOM = ByteOrderMark
    
    If BigEndian Then
        mIsBigEndian = True
        mCodePage = 12001
    End If
    
    If ThrowOnInvalidBytes Then
        Set mDecoderFallback = Statics.DecoderFallback.ExceptionFallback
        Set mEncoderFallback = Statics.EncoderFallback.ExceptionFallback
    End If
End Sub

Friend Sub InitClone(ByVal Other As UTF32Encoding, ByVal IsBigEndian As Boolean, ByVal EmitBOM As Boolean)
    Set mDecoderFallback = Other.DecoderFallback
    Set mEncoderFallback = Other.EncoderFallback
    mCodePage = Other.CodePage
    mIsBigEndian = IsBigEndian
    mEmitBOM = EmitBOM
    mIsReadOnly = False
End Sub

Private Sub Class_InitProperties()
    Set mDecoderFallback = Cor.NewDecoderReplacementFallback(ChrW$(&HFFFD))
    Set mEncoderFallback = Cor.NewEncoderReplacementFallback(ChrW$(&HFFFD))
    mCodePage = 12000
    mIsReadOnly = True
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Table() As EncodingTable
    Set Table = Statics.Encoding.EncodingTable
End Property

Friend Function GetByteCountCore(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByVal Encoder As UTF32Encoder) As Long
    GetByteCountCore = GetBytesCore(Chars, Index, Count, NullBytes, 0, Encoder)
End Function

Friend Function GetBytesCore(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal Encoder As UTF32Encoder) As Long
    Const BeforeSurrogateStart  As Long = SurrogateStart - 1
    Const AfterSurrogateEnd     As Long = SurrogateEnd + 1
    
    Dim FallbackBuffer  As EncoderFallbackBuffer
    Dim StartIndex      As Long
    Dim LastChar        As Long
    Dim Ch              As Long
    Dim BytesUBound     As Long
    Dim PutBytes        As Boolean
    
    If SAPtr(Bytes) <> vbNullPtr Then
        BytesUBound = UBound(Bytes)
        PutBytes = True
    End If
    
    If Not Encoder Is Nothing Then
        LastChar = Encoder.LastChar
        
        If PutBytes Then
            Encoder.LastChar = 0
        End If
    End If
    
    StartIndex = ByteIndex
    
    Do While CharCount > 0
        Ch = Chars(CharIndex) And &HFFFF&
        
        If LastChar = 0 Then
            Select Case Ch
                Case 0 To BeforeSurrogateStart, AfterSurrogateEnd To &HFFFF&
                    If PutBytes Then
                        If ByteIndex + vbSizeOfUTF32Char - 1 > BytesUBound Then _
                            ThrowEncoderSmallConversionBuffer
                        
                        ' This isn't quite a duplicate of the character assignment
                        ' later in this method. This version only assigns two bytes
                        ' from the character. The other assigns three bytes.
                        If mIsBigEndian Then
                            Bytes(ByteIndex) = 0
                            Bytes(ByteIndex + 1) = 0
                            Bytes(ByteIndex + 2) = (Ch \ vbShift8Bits) And &HFF
                            Bytes(ByteIndex + 3) = Ch And &HFF
                        Else
                            Bytes(ByteIndex) = Ch And &HFF
                            Bytes(ByteIndex + 1) = (Ch \ vbShift8Bits) And &HFF
                            Bytes(ByteIndex + 2) = 0
                            Bytes(ByteIndex + 3) = 0
                        End If
                    End If
                    
                    ByteIndex = ByteIndex + vbSizeOfUTF32Char
                Case HighSurrogateStart To HighSurrogateEnd
                    LastChar = Ch
                Case Else
                    ByteIndex = InsertFallbackCharacters(Ch, CharIndex, Bytes, ByteIndex, Encoder, FallbackBuffer)
            End Select
        Else
            Select Case Ch
                Case LowSurrogateStart To LowSurrogateEnd
                    If PutBytes Then
                        Dim Pair As Long
                        
                        If ByteIndex + vbSizeOfUTF32Char - 1 > BytesUBound Then _
                            ThrowEncoderSmallConversionBuffer
                        
                        Pair = (LastChar And &H3FF) * vbShift10Bits + (Ch And &H3FF) + UnicodePlane1Start
                        
                        ' This isn't quite a duplicate of the character assignment
                        ' earlier in this method. This version assigns three bytes
                        ' from the character. The other assigns two bytes.
                        If mIsBigEndian Then
                            Bytes(ByteIndex) = 0
                            Bytes(ByteIndex + 1) = (Pair \ vbShift16Bits) And &HFF
                            Bytes(ByteIndex + 2) = (Pair \ vbShift8Bits) And &HFF
                            Bytes(ByteIndex + 3) = Pair And &HFF
                        Else
                            Bytes(ByteIndex) = Pair And &HFF
                            Bytes(ByteIndex + 1) = (Pair \ vbShift8Bits) And &HFF
                            Bytes(ByteIndex + 2) = (Pair \ vbShift16Bits) And &HFF
                            Bytes(ByteIndex + 3) = 0
                        End If
                    End If
                    
                    ByteIndex = ByteIndex + 4
                Case Else
                    ' A high surrogate must be followed by a low surrogate or we need
                    ' to fallback the invalid sequence of characters.
                    ByteIndex = InsertFallbackCharacters(LastChar, CharIndex - 1, Bytes, ByteIndex, Encoder, FallbackBuffer)
                    CharIndex = CharIndex - 1
                    CharCount = CharCount + 1
            End Select
            
            LastChar = 0
        End If
                
        CharIndex = CharIndex + 1
        CharCount = CharCount - 1
    Loop
    
    If LastChar <> 0 Then
        If Encoder Is Nothing Then
            ByteIndex = InsertFallbackCharacters(LastChar, CharIndex - 1, Bytes, ByteIndex, Encoder, FallbackBuffer)
        ElseIf Encoder.Flushing Then
            ByteIndex = InsertFallbackCharacters(LastChar, CharIndex - 1, Bytes, ByteIndex, Encoder, FallbackBuffer)
        ElseIf PutBytes Then
            Encoder.LastChar = LastChar
        End If
    End If
    
    GetBytesCore = ByteIndex - StartIndex
End Function

Private Function InsertFallbackCharacters(ByVal Ch As Long, ByVal CharIndex As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal Encoder As UTF32Encoder, ByRef Fallback As EncoderFallbackBuffer) As Long
    Dim FallbackChar    As Integer
    Dim LastChar        As Integer
    
    If Fallback Is Nothing Then
        If Encoder Is Nothing Then
            Set Fallback = mEncoderFallback.CreateFallbackBuffer
        Else
            Set Fallback = Encoder.FallbackBuffer
        End If
    End If
    
    Fallback.Reset
    
    ' Didn't optimize the fallback because it's not really the happy path.
    If Fallback.Fallback(Ch, Index:=CharIndex) Then
        Do While Fallback.Remaining > 0
            FallbackChar = Fallback.GetNextChar
            
            If IsHighSurrogate(FallbackChar) Then
                If LastChar <> 0 Then _
                    Error.Argument Argument_InvalidCharSequenceNoIndex
                    
                LastChar = FallbackChar
            ElseIf IsLowSurrogate(FallbackChar) Then
                If LastChar = 0 Then _
                    Error.Argument Argument_InvalidCharSequenceNoIndex
                
                ByteIndex = InsertUTF32Char(ConvertToUtf32(LastChar, FallbackChar), Bytes, ByteIndex)
                LastChar = 0
            ElseIf LastChar <> 0 Then
                Error.Argument Argument_InvalidCharSequenceNoIndex
            Else
                ByteIndex = InsertUTF32Char(Char.ConvertToInt32(FallbackChar), Bytes, ByteIndex)
            End If
        Loop
    End If
    
    If LastChar <> 0 Then _
        Error.Argument Argument_InvalidCharSequenceNoIndex
        
    InsertFallbackCharacters = ByteIndex
End Function

Private Function InsertUTF32Char(ByVal UTF32Char As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    If SAPtr(Bytes) Then
        If ByteIndex + vbSizeOfUTF32Char - 1 > UBound(Bytes) Then _
            ThrowEncoderSmallConversionBuffer
    
        If mIsBigEndian Then
            UTF32Char = SwapEndian(UTF32Char)
        End If

        AsLong(Bytes(ByteIndex)) = UTF32Char
    End If
    
    InsertUTF32Char = ByteIndex + vbSizeOfUTF32Char
End Function

Friend Function GetCharCountCore(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, ByVal Decoder As UTF32Decoder) As Long
    GetCharCountCore = GetCharsCore(Bytes, Index, Count, NullChars, 0, Decoder)
End Function

Friend Function GetCharsCore(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal Decoder As UTF32Decoder) As Long
    Const BeforeSurrogateStart  As Long = SurrogateStart - 1
    Const AfterSurrogateEnd     As Long = SurrogateEnd + 1
    
    Dim PutChars        As Boolean
    Dim CharsUBound     As Long
    Dim BytesRead       As Long
    Dim UTF32Char       As Long
    Dim StartIndex      As Long
    Dim FallbackBuffer  As DecoderFallbackBuffer
    
    If SAPtr(Chars) Then
        CharsUBound = UBound(Chars)
        PutChars = True
    End If
    
    StartIndex = CharIndex
    
    If Not Decoder Is Nothing Then
        UTF32Char = Decoder.UTF32Char
        BytesRead = Decoder.BytesRead
        
        If PutChars Then
            Decoder.UTF32Char = 0
            Decoder.BytesRead = 0
        End If
    End If
    
    Do While ByteCount > 0
        If ByteCount > 3 And BytesRead = 0 Then
#If Release Then
            UTF32Char = Bytes(ByteIndex) Or _
                        Bytes(ByteIndex + 1) * vbShift8Bits Or _
                        Bytes(ByteIndex + 2) * vbShift16Bits Or _
                        Bytes(ByteIndex + 3) * vbShift24Bits
#Else
            UTF32Char = AsLong(Bytes(ByteIndex))
#End If

            BytesRead = 4
            ByteIndex = ByteIndex + 4
            ByteCount = ByteCount - 4
        Else
#If Release Then
            UTF32Char = ((UTF32Char \ vbShift8Bits) And &HFFFFFF) Or (Bytes(ByteIndex) * vbShift24Bits)
#Else
            UTF32Char = ShiftRightInt32(UTF32Char, 8) Or ShiftLeftInt32(Bytes(ByteIndex), 24)
#End If
        
            BytesRead = BytesRead + 1
            ByteIndex = ByteIndex + 1
            ByteCount = ByteCount - 1
        End If
        
        If BytesRead = 4 Then
            If mIsBigEndian Then
                UTF32Char = SwapEndian(UTF32Char)
            End If
                        
            Select Case UTF32Char
                ' Check if we have an normal character. These will be the most common characters to process.
                ' Characters are from U+0000 to U+D7FF and U+E000 to U+FFFF.
                Case 0 To BeforeSurrogateStart, AfterSurrogateEnd To &HFFFF&
                    If PutChars Then
                        If CharIndex > CharsUBound Then _
                            ThrowDecoderSmallConversionBuffer

    #If Release Then
                        Chars(CharIndex) = UTF32Char And &HFFFF&
    #Else
                        Chars(CharIndex) = AsWord(UTF32Char)
    #End If
                    End If

                    CharIndex = CharIndex + 1
                
                ' Check if we have an encoded high/low surrogate pair.
                ' Encoded surrogate pairs are from U+10000 to U+10FFFF.
                Case UnicodePlane1Start To UnicodePlane16End
                    If PutChars Then
                        If CharIndex + 1 > CharsUBound Then _
                            ThrowDecoderSmallConversionBuffer
    
                        UTF32Char = UTF32Char - UnicodePlane1Start
                        Chars(CharIndex) = ((UTF32Char \ vbShift10Bits) And &H3FF) + &HD800
                        Chars(CharIndex + 1) = (UTF32Char And &H3FF) + &HDC00
                    End If
    
                    CharIndex = CharIndex + 2
                    
                ' We have an invalid character, so fallback.
                Case Else
                    ' Swap the endian back because we are pulling the BytesUnknown array
                    ' from the character which had the endian swapped to make it little endian.
                    If mIsBigEndian Then
                        UTF32Char = SwapEndian(UTF32Char)
                    End If
    
                    CharIndex = InsertDecoderFallback(AsQByteArr(UTF32Char).Bytes, ByteIndex - vbSizeOfUTF32Char, Chars, CharIndex, Decoder, FallbackBuffer)
            End Select
            
            BytesRead = 0
            UTF32Char = 0
        End If
    Loop
    
    If BytesRead > 0 Then
        If CanFallback(Decoder) Then
            ' We didn't have enough bytes to decode another character
            ' so we will have to fallback the remaining bytes.
            Dim BytesUnknown() As Byte
            Dim i As Long
            
            ReDim BytesUnknown(0 To BytesRead - 1)
            UTF32Char = (UTF32Char \ ((4 - BytesRead) * vbShift8Bits)) And &HFFFFFF

            For i = 0 To BytesRead - 1
                BytesUnknown(i) = UTF32Char And &HFF&
                UTF32Char = (UTF32Char \ vbShift8Bits) And &HFFFFFF
            Next
    
            CharIndex = InsertDecoderFallback(BytesUnknown, ByteIndex - BytesRead, Chars, CharIndex, Decoder, FallbackBuffer)
        ElseIf PutChars Then
            Decoder.UTF32Char = UTF32Char
            Decoder.BytesRead = BytesRead
        End If
    End If
    
    GetCharsCore = CharIndex - StartIndex
End Function

Private Function CanFallback(ByVal Decoder As UTF32Decoder) As Boolean
    If Decoder Is Nothing Then
        CanFallback = True
    Else
        CanFallback = Decoder.Flushing
    End If
End Function

Private Function InsertDecoderFallback(ByRef BytesUnknown() As Byte, ByVal ByteIndex As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal Decoder As UTF32Decoder, ByRef FallbackBuffer As DecoderFallbackBuffer) As Long
    Dim Ch As Integer
    
    If FallbackBuffer Is Nothing Then
        If Decoder Is Nothing Then
            Set FallbackBuffer = mDecoderFallback.CreateFallbackBuffer
        Else
            Set FallbackBuffer = Decoder.FallbackBuffer
        End If
    End If
    
    FallbackBuffer.Reset
    
    If FallbackBuffer.Fallback(BytesUnknown, ByteIndex) Then
        If SAPtr(Chars) Then
            Do While FallbackBuffer.Remaining > 0
                Ch = FallbackBuffer.GetNextChar
                            
                If CharIndex > UBound(Chars) Then _
                    ThrowDecoderSmallConversionBuffer
            
                Chars(CharIndex) = Ch
                CharIndex = CharIndex + 1
            Loop
        Else
            CharIndex = CharIndex + FallbackBuffer.Remaining
        End If
    End If
    
    InsertDecoderFallback = CharIndex
End Function

Private Sub ThrowEncoderSmallConversionBuffer()
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowBytes, EncodingName, TypeName(EncoderFallback)), "Bytes")
End Sub

Private Sub ThrowDecoderSmallConversionBuffer()
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowChars, EncodingName, TypeName(DecoderFallback)), "Chars")
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

Private Function Encoding_GetByteCount(Chars As Variant, Optional Index As Variant, Optional Count As Variant) As Long
    Encoding_GetByteCount = GetByteCount(Chars, Index, Count)
End Function

Private Function Encoding_GetBytes(Chars As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    Encoding_GetBytes = GetBytes(Chars, Index, Count)
End Function

Private Function Encoding_GetBytesEx(Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Encoding_GetBytesEx = GetBytesEx(Chars, CharIndex, CharCount, Bytes, ByteIndex)
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

