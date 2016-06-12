VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASCIIEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: ASCIIEncoding
'

''
' Provides a set of methods used to encode and decode chars to and from bytes in ASCII format.
'
' @remarks
' Encoding characters to bytes using ASCII encoding will produce a 1:1 ratio for each
' character-to-byte encoding. If the character is greater than 127, then it will be
' replaced with character 63 (?).
' <p>To create a new <b>ASCIIEncoding</b> use the New function.
' <pre>
' Set en = New ASCIIEncoding
' </pre></p>
' @see Encoding
' @see UTF8Encoding
' @see UTF7Encoding
' @see UnicodeEncoding
'
Option Explicit
Implements IObject
Implements Encoding

Private Const QuestionMarkByte As Byte = 63

Private mPreamble() As Byte


''
' Returns if the current encoding uses single-byte code points.
'
' @return Returns True for single-byte, False for multi-byte code points.
Public Property Get IsSingleByte() As Boolean
    IsSingleByte = True
End Property

''
' Returns the encoding name to be used in with the mail agent body tags.
'
' @return The body tag compatible encoding name.
' @see Encoding
'
Public Property Get BodyName() As String
    BodyName = "us-ascii"
End Property

''
' Returns the code page identifier for this encoding.
'
' @return Code page identifier.
' @see Encoding
'
Public Property Get CodePage() As Long
    CodePage = 20127
End Property

''
' The name of this encoding.
'
' @return The encoding's name.
' @see Encoding
'
Public Property Get EncodingName() As String
    EncodingName = "US-ASCII"
End Property

''
' Returns the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
' @see Encoding
'
Public Property Get HeaderName() As String
    HeaderName = "us-ascii"
End Property

''
' Indicates if this encoding can be used by browsers to display text.
'
' @return The indication for browser display capable.
' @see Encoding
'
Public Property Get IsBrowserDisplay() As Boolean
    IsBrowserDisplay = False
End Property

''
' Indicates if this encoding can be used to save data with this encoding.
'
' @return The indication for the browser saving capable.
' @see Encoding
'
Public Property Get IsBrowserSave() As Boolean
    IsBrowserSave = False
End Property

''
' Indicates if this encoding can be used to display mail and news by
' mail and news clients.
'
' @return Indication for mail and news client capable.
' @see Encoding
'
Public Property Get IsMailNewsDisplay() As Boolean
    IsMailNewsDisplay = True
End Property

''
' Indicates if this encoding can be used to save date by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
' @see Encoding
'
Public Property Get IsMailNewsSave() As Boolean
    IsMailNewsSave = True
End Property

''
' Returns the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
' @see Encoding
'
Public Property Get WebName() As String
    WebName = "us-ascii"
End Property

''
' Returns the Windows Operating System's code page for this encoding.
'
' @return The Windows code page for this encoding.
' @see Encoding
'
Public Property Get WindowsCodePage() As Long
    WindowsCodePage = 1252
End Property

''
' Returns the number of bytes that would be produced from the set of characters using this encoding.
'
' @param Chars A set of characters used to calculated the number of bytes once encoded with this encoding.
' @param Index The starting index of the first character to be used.
' @param Count The number of characters to include in the calculation.
' @return The number of bytes that would be produced by encoding the set of characters.
' @remarks The <i>Chars</i> parameter will accept either a String or an Integer array.
' <p>If a string is passed in, then the Index and Count parameters are ignored.</p>
' @see Encoding
'
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Select Case VarType(Chars)
        Case vbString
            GetByteCount = GetByteCountString(Chars)
        Case corIntegerArray
            GetByteCount = GetByteCountChars(Chars, Index, Count)
        Case Else
            Error.Argument Argument_ArrayRequired, "Chars"
    End Select
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
' @see Encoding
'
Public Function GetBytes(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    Dim AttachedChars() As Integer
    AttachedChars = AsChars(Chars)
    
    On Error GoTo Catch
    
    If SAPtr(AttachedChars) = vbNullPtr Then _
        Error.ArgumentNull "Chars", ArgumentNull_Array
    
    Dim Range As ListRange
    Range = OptionalRange(Index, Count, LBound(AttachedChars), CorArray.LengthFirstDim(AttachedChars))

    If Range.Index < LBound(AttachedChars) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count - 1 > UBound(AttachedChars) Then _
        Error.ArgumentOutOfRange "Chars", ArgumentOutOfRange_IndexCount

    Dim Result() As Byte
    If Range.Count > 0 Then
        ReDim Result(0 To GetByteCount(AttachedChars, Range.Index, Range.Count) - 1)
        GetBytesCore AttachedChars, Range.Index, Range.Count, Result, 0
        GetBytes = Result
    Else
        GetBytes = Cor.NewBytes()
    End If

    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars AttachedChars
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
' @see Encoding
'
Public Function GetBytesEx(ByRef Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Dim AttachedChars() As Integer
    AttachedChars = AsChars(Chars)
    
    On Error GoTo Catch
    
    If SAPtr(AttachedChars) = vbNullPtr Then _
        Error.ArgumentNull "Chars", ArgumentNull_Array
    If SAPtr(Bytes) = vbNullPtr Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    If CharIndex < LBound(AttachedChars) Then _
        Error.ArgumentOutOfRange "CharIndex", ArgumentOutOfRange_LBound
    If CharCount < 0 Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_NeedNonNegNum
    If CharIndex + CharCount - 1 > UBound(AttachedChars) Then _
        Error.ArgumentOutOfRange "Chars", ArgumentOutOfRange_IndexCountBuffer
    If ByteIndex < LBound(Bytes) Or ByteIndex > UBound(Bytes) Then _
        Error.ArgumentOutOfRange "ByteIndex", ArgumentOutOfRange_Index
    If ByteIndex + CharCount - 1 > UBound(Bytes) Then _
        Error.Argument Argument_SmallConversionBuffer, "Bytes"
    
    GetBytesEx = GetBytesCore(AttachedChars, CharIndex, CharCount, Bytes, ByteIndex)
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars AttachedChars
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
'
Public Function GetCharCount(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    CheckValidSingleDimArray Bytes, Parameter_Bytes
    
    Dim Range As ListRange
    Range = OptionalRange(Index, Count, LBound(Bytes), CorArray.LengthFirstDim(Bytes))

    If Range.Index < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count - 1 > UBound(Bytes) Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
    
    GetCharCount = Range.Count
End Function

''
' Decodes a set of bytes into a set of characters.
'
' @param Bytes The set of bytes to be decoded into characters.
' @param ByteIndex The index of the first byte to be used in decoding.
' @param ByteCount The number of bytes to be used in decoding.
' @return An array of characters decoded from the set of bytes.
' @see Encoding
'
Public Function GetChars(ByRef Bytes() As Byte, Optional ByRef ByteIndex As Variant, Optional ByRef ByteCount As Variant) As Integer()
    CheckValidSingleDimArray Bytes, Parameter_Bytes

    Dim Range As ListRange
    Range = OptionalRange(ByteIndex, ByteCount, LBound(Bytes), UBound(Bytes) - LBound(Bytes) + 1, Parameter_ByteIndex, Parameter_ByteCount)

    If Range.Index < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "ByteIndex", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count - 1 > UBound(Bytes) Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
    
    Dim Chars() As Integer
    If Range.Count > 0 Then
        ReDim Chars(0 To GetCharCount(Bytes, Range.Index, Range.Count) - 1)
        GetCharsCore Bytes, Range.Index, Range.Count, Chars, 0
    Else
        Chars = Cor.NewIntegers()
    End If
    
    GetChars = Chars
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
' @see Encoding
'
Public Function GetCharsEx(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    CheckValidSingleDimArray Bytes, Parameter_Bytes
    CheckValidSingleDimArray Chars, Parameter_Chars
    
    If ByteIndex < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "ByteIndex", ArgumentOutOfRange_LBound
    If ByteCount < 0 Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_NeedNonNegNum
    If ByteIndex + ByteCount - 1 > UBound(Bytes) Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
    If CharIndex < LBound(Chars) Then _
        Error.ArgumentOutOfRange "CharIndex", ArgumentOutOfRange_LBound
    If CharIndex + ByteCount - 1 > UBound(Chars) Then _
        Error.Argument Argument_SmallConversionBuffer, "Chars"
        
    GetCharsEx = GetCharsCore(Bytes, ByteIndex, ByteCount, Chars, CharIndex)
End Function

''
' Returns the maximum number of characters than can be decoded from the number of bytes specified.
'
' @param ByteCount The number of bytes to be decoded.
' @return The maximum number of characters that can be decoded from the specified number of bytes.
' @see Encoding
'
Public Function GetMaxCharCount(ByVal ByteCount As Long) As Long
    If ByteCount < 0 Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_NeedNonNegNum
        
    GetMaxCharCount = ByteCount
End Function

''
' Returns the maximum number of bytes that can be created from a specific number of characters.
'
' @param CharCount The number of characters to be encoded.
' @return The maximum number of bytes that can be generated from the specified number of characters.
' @see Encoding
'
Public Function GetMaxByteCount(ByVal CharCount As Long) As Long
    If CharCount < 0 Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_NeedNonNegNum
        
    GetMaxByteCount = CharCount
End Function

''
' Returns an encoder that maintains state.
'
' @return The stateful encoder.
' @remarks Since encoding Unicode characters equates to exactly 2 bytes per
' character, there is no state that is maintained between calls. This method
' functions identically to GetBytes.
' @see Encoding
'
Public Function GetEncoder() As Encoder
    Dim Encoder As New DefaultEncoder
    Encoder.Init Me
    Set GetEncoder = Encoder
End Function

''
' Returns a decoder that maintains state.
'
' @return The stateful decoder.
' @remarks Unlike UnicodeEncoding, the decoder maintains state between decoding calls.
' Since a single unicode character required exactly 2 bytes to be decoded, it is possible
' that an odd number of bytes may be attempted to be decoding, creating an orphaned byte
' which represents only half of a unicode character. The remaining byte is held until
' the next decoding call, and the byte is then inserted at the beginning of the next set
' of bytes to be decoded, picking up byte decoding where the orphaned byte left off.
' @see Encoding
'
Public Function GetDecoder() As Decoder
    Dim Decoder As New DefaultDecoder
    Decoder.Init Me
    Set GetDecoder = Decoder
End Function

''
' Returns an array of bytes that represents this encoding.
'
' @return A byte array containg the bytes to be used to identify this encoding type.
' @remarks <p>If the encoding was created setting the <i>BigEndian</i> paramter
' to true, then this will return the Unicode byte array identifer &HFEFF otherwise
' &hFFFE for a non-bigEndian encoding. If no identifier is to be emitted, then an
' empty byte array is returned (not a null array).
' @see Encoding
'
Public Function GetPreamble() As Byte()
    If SAPtr(mPreamble) = vbNullPtr Then
        mPreamble = Cor.NewBytes()
    End If
    
    GetPreamble = mPreamble
End Function

''
' Decodes a set of bytes into a String.
'
' @param Bytes The set of bytes to be decoded into a string.
' @param ByteIndex The index of the first byte to be decoded.
' @param ByteCount The number of bytes to be used in the decoding.
' @return A string containing the decoded set of bytes.
' @see Encoding
'
Public Function GetString(ByRef Bytes() As Byte, Optional ByRef ByteIndex As Variant, Optional ByRef ByteCount As Variant) As String
    CheckValidSingleDimArray Bytes, Parameter_Bytes

    Dim Range As ListRange
    Range = OptionalRange(ByteIndex, ByteCount, LBound(Bytes), UBound(Bytes) - LBound(Bytes) + 1, Parameter_ByteIndex, Parameter_ByteCount)

    If Range.Index < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "ByteIndex", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count - 1 > UBound(Bytes) Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer

    GetString = SysAllocStringLen(0, Range.Count)
    
    Dim Chars() As Integer
    Chars = AllocChars(GetString)
    GetCharsCore Bytes, Range.Index, Range.Count, Chars, 0
    FreeChars Chars
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
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
' @see IObject
'
Public Function GetHashCode() As Long
    GetHashCode = CodePage
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetByteCountString(ByRef Chars As Variant) As Long
    GetByteCountString = Len(Chars)
End Function

Private Function GetByteCountChars(ByRef Chars As Variant, ByRef Index As Variant, ByRef Count As Variant) As Long
    CheckValidSingleDimArray Chars, Parameter_Chars
        
    Dim Range As ListRange
    Range = OptionalRange(Index, Count, LBound(Chars), CorArray.LengthFirstDim(Chars))

    If Range.Index < LBound(Chars) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count - 1 > UBound(Chars) Then _
        Error.ArgumentOutOfRange "Chars", ArgumentOutOfRange_IndexCountBuffer

    GetByteCountChars = Range.Count
End Function

Private Function GetBytesCore(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Dim i   As Long
    Dim Ch  As Long
    For i = 0 To Count - 1
        Ch = Chars(Index + i)
        Select Case Ch
            ' we need to have a range because by using Integers to
            ' represent unicode characters, we could end up with
            ' negative numbers.
            Case 0 To 127
                Bytes(ByteIndex + i) = Ch
            Case Else
                Bytes(ByteIndex + i) = QuestionMarkByte
        End Select
    Next
    GetBytesCore = Count
End Function

Private Function GetCharsCore(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim i As Long
    Dim b As Byte
    For i = 0 To ByteCount - 1
        b = Bytes(ByteIndex + i)
        If b < 128 Then
            Chars(CharIndex + i) = b
        Else
            Chars(CharIndex + i) = QuestionMarkByte
        End If
    Next i
    GetCharsCore = ByteCount
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


