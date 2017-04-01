VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnicodeEncoding"
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
' Module: UnicodeEncoding
'

''
' A set of functions used to convert unicode characters to and from byte arrays.
'
' @remarks <p>This encoding can encode characters using big-endian or little-endian
' notation. The default is little-endian (least significant byte first).</p>
' <p>A Byte order Mark will be emitted if selected using the Cor.NewUnicodeEncoding constructor.</p>
'
' @see Constructors
' @see Encoding
'
Option Explicit
Implements IObject
Implements Encoding
Implements ICloneable

Private mIsReadOnly         As Boolean
Private mIsBigEndian        As Boolean
Private mEmitBOM            As Boolean
Private mCodePage           As Long
Private mDecoderFallback    As DecoderFallback
Private mEncoderFallback    As EncoderFallback


''
' Creates a clone of the current Encoding instance.
'
' @return A cloned instance of a UnicodeEncoding object.
' @remarks A cloned instance will have the read-only limitations removed.
'
Public Function Clone() As UnicodeEncoding
    Set Clone = New UnicodeEncoding
    Clone.InitClone Me, mIsBigEndian, mEmitBOM
End Function

''
' Returns if the current encoding uses single-byte code points.
'
' @return Returns True for single-byte, False for multi-byte code points.
Public Property Get IsSingleByte() As Boolean
    IsSingleByte = False
End Property

''
' Returns the encoding name to be used in with the mail agent body tags.
'
' @return The body tag compatible encoding name.
'
Public Property Get BodyName() As String
    BodyName = "utf-16"
End Property

''
' Returns the code page identifier for this encoding.
'
' @return Code page identifier.
'
Public Property Get CodePage() As Long
    CodePage = mCodePage
End Property

''
' The name of this encoding.
'
' @return The encoding's name.
'
Public Property Get EncodingName() As String
    EncodingName = "Unicode"
End Property

''
' Returns the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
'
Public Property Get HeaderName() As String
    HeaderName = "utf-16"
End Property

''
' Indicates if this encoding can be used by browsers to display text.
'
' @return The indication for browser display capable.
'
Public Property Get IsBrowserDisplay() As Boolean
    IsBrowserDisplay = False
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
    IsMailNewsDisplay = False
End Property

''
' Indicates if this encoding can be used to save date by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
'
Public Property Get IsMailNewsSave() As Boolean
    IsMailNewsSave = False
End Property

''
' Returns the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
'
Public Property Get WebName() As String
    WebName = "utf-16"
End Property

''
' Returns the Windows Operating System's code page for this encoding.
'
' @return The Windows code page for this encoding.
'
Public Property Get WindowsCodePage() As Long
    WindowsCodePage = mCodePage
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
' is zero based for both the String and Integer array, not 1 based like Mid$.
'
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim LocalChars()    As Integer
    Dim Range           As ListRange
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    If SAPtr(LocalChars) = vbNullPtr Then _
        Error.ArgumentNull "Chars", ArgumentNull_Array
        
    Range = GetArrayRange(LocalChars, Index, Count)
    
    If Range.Index < LBound(LocalChars) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayLB
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If UBound(LocalChars) - Range.Index + 1 < Range.Count Then _
        Error.ArgumentOutOfRange "Chars", ArgumentOutOfRange_IndexCountBuffer
    
    GetByteCount = GetByteCountCore(LocalChars, Range.Index, Range.Count)
    
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
    Dim Output()        As Byte
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    If SAPtr(LocalChars) = vbNullPtr Then _
        Error.ArgumentNull "Chars", ArgumentNull_Array
    
    Range = GetArrayRange(LocalChars, Index, Count)
    
    If Range.Index < LBound(LocalChars) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayLB
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If UBound(LocalChars) - Range.Index + 1 < Range.Count Then _
        Error.ArgumentOutOfRange "Chars", ArgumentOutOfRange_IndexCountBuffer
    
    If Range.Count > 0 Then
        ReDim Output(0 To GetByteCountCore(LocalChars, Range.Index, Range.Count) - 1)
        GetBytesCore LocalChars, Range.Index, Range.Count, Output, 0
    Else
        Output = Cor.NewBytes()
    End If
    
    GoSub Finally
    GetBytes = Output
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
    Dim LocalChars() As Integer
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    If SAPtr(LocalChars) = vbNullPtr Or SAPtr(Bytes) = vbNullPtr Then _
        Error.ArgumentNull IIf(SAPtr(LocalChars) = vbNullPtr, "Chars", "Bytes"), ArgumentNull_Array
    If CharIndex < LBound(LocalChars) Then _
        Error.ArgumentOutOfRange "CharIndex", ArgumentOutOfRange_ArrayLB
    If ByteIndex < LBound(Bytes) Or (ByteIndex > UBound(Bytes) And Len1D(Bytes) > 0) Then _
        Error.ArgumentOutOfRange "ByteIndex", ArgumentOutOfRange_Index
    If CharCount < 0 Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_NeedNonNegNum
    If UBound(LocalChars) - CharIndex + 1 < CharCount Then _
        Error.ArgumentOutOfRange "Chars", ArgumentOutOfRange_IndexCountBuffer
    
    If CharCount > 0 Then
        GetBytesEx = GetBytesCore(LocalChars, CharIndex, CharCount, Bytes, ByteIndex)
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
'
Public Function GetCharCount(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim Range As ListRange
    
    If SAPtr(Bytes) = vbNullPtr Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    
    Range = GetArrayRange(Bytes, Index, Count)
    
    If Range.Index < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayLB
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If UBound(Bytes) - Range.Index + 1 < Range.Count Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
    
    If Range.Count > 0 Then
        GetCharCount = GetCharCountCore(Bytes, Range.Index, Range.Count, Nothing)
    End If
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
    
    If SAPtr(Bytes) = vbNullPtr Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    
    Range = GetArrayRange(Bytes, Index, Count)
    
    If Range.Index < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayLB
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If UBound(Bytes) - Range.Index + 1 < Range.Count Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
    
    If Range.Count > 0 Then
        ReDim Output(0 To GetCharCountCore(Bytes, Range.Index, Range.Count, Nothing) - 1)
        GetCharsCore Bytes, Range.Index, Range.Count, Output, 0
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
    If SAPtr(Bytes) = vbNullPtr Or SAPtr(Chars) = vbNullPtr Then _
        Error.ArgumentNull IIf(SAPtr(Bytes) = vbNullPtr, "Bytes", "Chars"), ArgumentNull_Array
    If ByteIndex < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "ByteIndex", ArgumentOutOfRange_ArrayLB
    If CharIndex < LBound(Chars) Or (CharIndex > UBound(Chars) And Len1D(Chars) > 0) Then _
        Error.ArgumentOutOfRange "CharIndex", ArgumentOutOfRange_Index
    If ByteCount < 0 Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_NeedNonNegNum
    If UBound(Bytes) - ByteIndex + 1 < ByteCount Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
    
    If ByteCount > 0 Then
        GetCharsEx = GetCharsCore(Bytes, ByteIndex, ByteCount, Chars, CharIndex)
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

    Dim ByteCount As Currency
    ByteCount = 1@ + CharCount
    ByteCount = ByteCount * 2
    
    If ByteCount > &H7FFFFFFF Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_GetByteCountOverflow
        
    GetMaxByteCount = ByteCount
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

    Dim CharCount As Currency
    CharCount = CCur(CLng(ByteCount \ 2)) + (ByteCount And 1) + 1
    
    If DecoderFallback.MaxCharCount > 1 Then
        CharCount = CharCount * DecoderFallback.MaxCharCount
    End If
    
    If CharCount > &H7FFFFFFF Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_GetCharCountOverflow
        
    GetMaxCharCount = CharCount
End Function

''
' Returns an array of bytes that represents this encoding.
'
' @return A byte array containg the bytes to be used to identify this encoding type.
' @remarks <p>If the encoding was created setting the <i>BigEndian</i> paramter
' to true, then this will return the Unicode byte array identifer &HFEFF otherwise
' &hFFFE for a non-bigEndian encoding. If no identifier is to be emitted, then an
' empty byte array is returned (not a null array).
'
Public Function GetPreamble() As Byte()
    If mEmitBOM Then
        If mIsBigEndian Then
            GetPreamble = Cor.NewBytes(&HFE, &HFF)
        Else
            GetPreamble = Cor.NewBytes(&HFF, &HFE)
        End If
    Else
        GetPreamble = Cor.NewBytes()
    End If
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
    Dim Range   As ListRange
    Dim Chars() As Integer
    Dim Length  As Long
    
    If SAPtr(Bytes) = vbNullPtr Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
    
    Range = GetArrayRange(Bytes, Index, Count)
    
    If Range.Index < LBound(Bytes) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayLB
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If UBound(Bytes) - Range.Index + 1 < Range.Count Then _
        Error.ArgumentOutOfRange "Bytes", ArgumentOutOfRange_IndexCountBuffer
    
    GetString = String$(GetCharCountCore(Bytes, Range.Index, Range.Count, Nothing), 0)
    Chars = AllocChars(GetString)
    
    On Error GoTo Catch
    Length = GetCharsCore(Bytes, Range.Index, Range.Count, Chars, 0)
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars Chars
    Return
End Function

''
' Returns an encoder that maintains state.
'
' @return The stateful encoder.
' @remarks Since encoding Unicode characters equates to exactly 2 bytes per
' character, there is no state that is maintained between calls. This method
' functions identically to GetBytes.
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
'
Public Function GetDecoder() As Decoder
    Dim Decoder As New UnicodeDecoder
    Decoder.Init Me
    Set GetDecoder = Decoder
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
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
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Dim Other As UnicodeEncoding
    
    On Error GoTo Catch
    Set Other = Value
    Equals = (Other.GetHashCode = GetHashCode)
    
Catch:
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = mCodePage

    If mIsBigEndian Then
        GetHashCode = GetHashCode Or &H80000000
    End If
    
    If mEmitBOM Then
        GetHashCode = GetHashCode Or &H40000000
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal BigEndian As Boolean, ByVal ByteOrderMark As Boolean, ByVal ThrowOnInvalidBytes As Boolean)
    mIsBigEndian = BigEndian
    mEmitBOM = ByteOrderMark
    
    If BigEndian Then
        mCodePage = 1201
    End If
    
    If ThrowOnInvalidBytes Then
        Set mDecoderFallback = Statics.DecoderFallback.ExceptionFallback
        Set mEncoderFallback = Statics.EncoderFallback.ExceptionFallback
    End If
End Sub

Friend Sub InitClone(ByVal Other As UnicodeEncoding, ByVal IsBigEndian As Boolean, ByVal EmitBOM As Boolean)
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
    mCodePage = 1200
    mIsReadOnly = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBytesCore(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Dim Ch          As Long
    Dim ByteStart   As Long
    Dim PrevChar    As Long
    Dim Fallback    As EncoderFallbackBuffer
    
    If CharCount = 0 Then
        Exit Function
    End If
    
    ByteStart = ByteIndex
    
    Do While CharCount > 0
        Ch = Chars(CharIndex) And &HFFFF&
        
        If IsSurrogate(Ch) Then
            If IsHighSurrogate(Ch) Then
                If PrevChar <> 0 Then
                    ' we have two high surrogates in a row
                    ByteIndex = InsertEncoderFallback(PrevChar, Chars, CharIndex, Bytes, ByteIndex, Fallback)
                End If
                
                PrevChar = Ch
            ElseIf PrevChar <> 0 Then
                ' we have valid surrogate pair
                ByteIndex = InsertChar(PrevChar, Bytes, ByteIndex)
                ByteIndex = InsertChar(Ch, Bytes, ByteIndex)
                PrevChar = 0
            Else
                ' low surrogate without previous high surrogate
                ByteIndex = InsertEncoderFallback(Ch, Chars, CharIndex, Bytes, ByteIndex, Fallback)
            End If
        Else
            ByteIndex = InsertChar(Ch, Bytes, ByteIndex)
        End If
        
        CharIndex = CharIndex + 1
        CharCount = CharCount - 1
    Loop
    
    If PrevChar <> 0 Then
        ' we have a hanging high surrogate
        ByteIndex = InsertEncoderFallback(PrevChar, Chars, CharIndex, Bytes, ByteIndex, Fallback)
    End If
    
    GetBytesCore = ByteIndex - ByteStart
End Function

Private Function InsertEncoderFallback(ByVal Ch As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByRef Fallback As EncoderFallbackBuffer)
    If Fallback Is Nothing Then
        Set Fallback = mEncoderFallback.CreateFallbackBuffer
    End If
    
    Fallback.Reset
    
    If IsHighSurrogate(Ch) Then
        If CharIndex < UBound(Chars) Then
            Dim NextCh As Integer
            NextCh = Chars(CharIndex + 1)
            
            If IsLowSurrogate(NextCh) Then
                If Fallback.Fallback(Ch, NextCh) Then
                    Do While Fallback.Remaining > 0
                        ByteIndex = InsertChar(Fallback.GetNextChar, Bytes, ByteIndex)
                    Loop
                    
                    InsertEncoderFallback = ByteIndex
                    Exit Function
                End If
            End If
        End If
    End If
    
    If Fallback.Fallback(Ch) Then
        Do While Fallback.Remaining > 0
            ByteIndex = InsertChar(Fallback.GetNextChar, Bytes, ByteIndex)
        Loop
        
        InsertEncoderFallback = ByteIndex
    End If
End Function

Private Function InsertChar(ByVal Ch As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    If ByteIndex + 1 > UBound(Bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowBytes, EncodingName, TypeName(EncoderFallback)), "Bytes")
    
    If Not mIsBigEndian Then
        Bytes(ByteIndex) = Ch And &HFF&
        Bytes(ByteIndex + 1) = (Ch And &HFF00&) \ vbShift8Bits
    Else
        Bytes(ByteIndex) = (Ch And &HFF00&) \ vbShift8Bits
        Bytes(ByteIndex + 1) = Ch And &HFF&
    End If
    
    InsertChar = ByteIndex + 2
End Function

Friend Function GetCharCountCore(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long, ByVal Decoder As UnicodeDecoder) As Long
    Dim CharCount       As Long
    Dim ByteEnd         As Long
    Dim LastByte        As Long
    Dim LastChar        As Long
    Dim FallbackBuffer  As DecoderFallbackBuffer
    
    ByteEnd = Index + Count
    
    If Decoder Is Nothing Then
        LastByte = -1
    Else
        LastByte = Decoder.LastByte
        LastChar = Decoder.LastChar
    End If
    
    Do While Index < ByteEnd
        If LastByte < 0 Then
            LastByte = Bytes(Index)
            Index = Index + 1
            
            If Index >= ByteEnd Then
                Exit Do
            End If
        End If
        
        Dim Ch As Long
        If mIsBigEndian Then
            Ch = (LastByte * vbShift8Bits) Or Bytes(Index)
        Else
            Ch = (Bytes(Index) * vbShift8Bits) Or LastByte
        End If
        
        Index = Index + 1
        LastByte = -1
        
        If IsSurrogate(Ch) Then
            If IsHighSurrogate(Ch) Then
                If LastChar > 0 Then
                    ' we have two high surrogates in a row, so add fallback count for
                    ' first surrogate only. We may still have a low surrogate to follow.
                    CharCount = CharCount + GetDecoderFallbackCount(LastChar, FallbackBuffer, Decoder)
                End If
                
                LastChar = Ch
            ElseIf LastChar = 0 Then
                ' we have a low surrogate with a missing previous high surrogate
                ' so add fallback count.
                CharCount = CharCount + GetDecoderFallbackCount(Ch, FallbackBuffer, Decoder)
            Else
                ' we have a valid surrogate pair so count two characters.
                CharCount = CharCount + 2
                LastChar = 0
            End If
        ElseIf LastChar > 0 Then
            ' we have a regular character leaving a hanging high surrogate, so add fallback count
            ' and count the regular character.
            CharCount = CharCount + 1 + GetDecoderFallbackCount(LastChar, FallbackBuffer, Decoder)
            LastChar = 0
        Else
            ' just a regular character
            CharCount = CharCount + 1
        End If
    Loop
    
    ' we could wrap all this with LastChar > 0 statement, but that
    ' would create a bug by not clearing Decoder.LastChar when LastChar = 0.
    If Decoder Is Nothing Then
        If LastChar > 0 Then
            CharCount = CharCount + GetDecoderFallbackCount(LastChar, FallbackBuffer, Decoder)
        End If
    ElseIf Decoder.Flushing Then
        If LastChar > 0 Then
            CharCount = CharCount + GetDecoderFallbackCount(LastChar, FallbackBuffer, Decoder)
        End If
    Else
        Decoder.LastChar = LastChar
    End If
    
    ' similar to dealing with LastChar, we can't wrap all this with LastByte >= 0
    ' because it will cause a similar bug of not clearing the Decoder.LastByte.
    If Decoder Is Nothing Then
        If LastByte >= 0 Then
            CharCount = CharCount + GetDecoderFallbackCount(LastByte, FallbackBuffer, Decoder)
        End If
    ElseIf Decoder.Flushing Then
        If LastByte >= 0 Then
            CharCount = CharCount + GetDecoderFallbackCount(LastByte, FallbackBuffer, Decoder)
        End If
    Else
        Decoder.LastByte = LastByte
    End If
    
    GetCharCountCore = CharCount
End Function

Private Function GetDecoderFallbackCount(ByVal Ch As Long, ByRef FallbackBuffer As DecoderFallbackBuffer, ByVal Decoder As Decoder) As Long
    Dim ByteBuffer(0 To 1) As Byte
    
    If mIsBigEndian Then
        ByteBuffer(0) = (Ch \ vbShift8Bits) And &HFF
        ByteBuffer(1) = Ch And &HFF
    Else
        ByteBuffer(0) = Ch And &HFF
        ByteBuffer(1) = (Ch \ vbShift8Bits) And &HFF
    End If
    
    If FallbackBuffer Is Nothing Then
        If Decoder Is Nothing Then
            Set FallbackBuffer = mDecoderFallback.CreateFallbackBuffer
        Else
            Set FallbackBuffer = Decoder.Fallback.CreateFallbackBuffer
        End If
    End If
    
    If FallbackBuffer.Fallback(ByteBuffer, 0) Then
        GetDecoderFallbackCount = FallbackBuffer.Remaining
        FallbackBuffer.Reset
    End If
End Function

Private Function GetByteCountCore(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    Dim ByteCount   As Long
    Dim PrevChar    As Integer
    Dim Fallback    As EncoderFallbackBuffer
    Dim MaxIndex    As Long
    
    MaxIndex = Index + Count
    
    Do While Index < MaxIndex
        Dim Ch As Integer
        Ch = Chars(Index)
        
        If IsSurrogate(Ch) Then
            If IsHighSurrogate(Ch) Then
                If PrevChar <> 0 Then
                    ' got two high surrogates in a row
                    ByteCount = ByteCount + GetEncoderFallbackCount(PrevChar, Chars, Index, Fallback)
                End If
                
                PrevChar = Ch
            ElseIf PrevChar <> 0 Then
                ' got a valid surrogate pair
                ByteCount = ByteCount + vbSizeOfChar * 2
                PrevChar = 0
            Else
                ' low surrogate without previous high surrogate
                ByteCount = ByteCount + GetEncoderFallbackCount(PrevChar, Chars, Index, Fallback)
            End If
        ElseIf PrevChar <> 0 Then
            ' got a previous high surrogate with no low surrogate
            ByteCount = ByteCount + vbSizeOfChar + GetEncoderFallbackCount(PrevChar, Chars, Index, Fallback)
            PrevChar = 0
        Else
            ByteCount = ByteCount + vbSizeOfChar
        End If
        
        Index = Index + 1
    Loop
    
    If PrevChar <> 0 Then
        ' we have a hanging high surrogate
        ByteCount = ByteCount + GetEncoderFallbackCount(PrevChar, Chars, Index, Fallback)
    End If
    
    GetByteCountCore = ByteCount
End Function

Private Function GetEncoderFallbackCount(ByVal Ch As Integer, ByRef Chars() As Integer, ByVal CharIndex As Long, ByRef Fallback As EncoderFallbackBuffer) As Long
    If Fallback Is Nothing Then
        Set Fallback = mEncoderFallback.CreateFallbackBuffer
    End If
    
    Fallback.Reset
    
    If IsHighSurrogate(Ch) Then
        If CharIndex < UBound(Chars) Then
            Dim NextCh As Integer
            NextCh = Chars(CharIndex + 1)
            
            If IsLowSurrogate(NextCh) Then
                If Fallback.Fallback(Ch, NextCh) Then
                    GetEncoderFallbackCount = Fallback.Remaining * vbSizeOfChar
                    Exit Function
                End If
            End If
        End If
    End If
    
    If Fallback.Fallback(Ch) Then
        GetEncoderFallbackCount = Fallback.Remaining * vbSizeOfChar
    End If
End Function

Private Function GetCharsCore(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim ByteEnd         As Long
    Dim CharStart       As Long
    Dim CharEnd         As Long
    Dim LastByte        As Long
    Dim LastChar        As Long
    Dim FallbackBuffer  As DecoderFallbackBuffer
    
    ByteEnd = ByteIndex + ByteCount
    CharStart = CharIndex
    CharEnd = UBound(Chars)
    LastByte = -1
    
    Do While ByteIndex < ByteEnd
        If LastByte < 0 Then
            LastByte = Bytes(ByteIndex)
            ByteIndex = ByteIndex + 1
        Else
            Dim Ch As Long
            If mIsBigEndian Then
                Ch = (LastByte * vbShift8Bits) Or Bytes(ByteIndex)
            Else
                Ch = (Bytes(ByteIndex) * vbShift8Bits) Or LastByte
            End If
            
            ByteIndex = ByteIndex + 1
            LastByte = -1
            
            If IsSurrogate(Ch) Then
                If IsHighSurrogate(Ch) Then
                    ' we have preceeding high surrogate followed
                    ' by this high surrogate.
                    If LastChar > 0 Then
                        InsertDecoderFallback LastChar, Chars, CharIndex, FallbackBuffer
                    End If
                    
                    LastChar = Ch
                    GoTo Continue
                ElseIf LastChar = 0 Then
                    ' we have a low surrogate and expected the preceeding
                    ' character to be a high surrogate.
                    InsertDecoderFallback Ch, Chars, CharIndex, FallbackBuffer
                    GoTo Continue
                Else
                    ' we have a valid surrogate pair.
                    If CharIndex > CharEnd Then
                        ' handle small char buffer
                    End If
                    
                    Chars(CharIndex) = AsWord(LastChar)
                    CharIndex = CharIndex + 1
                    LastChar = 0
                End If
            ElseIf LastChar > 0 Then
                ' we have a preceeding high surrogate, but was not
                ' followed by a low surrogate.
                InsertDecoderFallback LastChar, Chars, CharIndex, FallbackBuffer
                LastChar = 0
            End If
            
            If CharIndex > CharEnd Then
                ' handle small char buffer
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowChars, EncodingName, TypeName(DecoderFallback)), "Chars")
            End If
            
            Chars(CharIndex) = AsWord(Ch)
            CharIndex = CharIndex + 1
        End If
        
Continue:
    Loop
    
    If LastChar > 0 Then
        ' have an unmatched high surrogate, handle fallback
        InsertDecoderFallback LastChar, Chars, CharIndex, FallbackBuffer
    End If
    
    If LastByte >= 0 Then
        ' have a partial char, handle fallback
        InsertDecoderFallback LastByte, Chars, CharIndex, FallbackBuffer
    End If
    
    GetCharsCore = CharIndex - CharStart
End Function

Private Sub InsertDecoderFallback(ByVal UnknownChar As Long, ByRef Chars() As Integer, ByRef CharIndex As Long, ByRef FallbackBuffer As DecoderFallbackBuffer)
    Dim ByteBuffer(0 To 1) As Byte
    
    If mIsBigEndian Then
        ByteBuffer(0) = (UnknownChar \ vbShift8Bits) And &HFF
        ByteBuffer(1) = UnknownChar And &HFF
    Else
        ByteBuffer(0) = UnknownChar And &HFF
        ByteBuffer(1) = (UnknownChar \ vbShift8Bits) And &HFF
    End If
    
    If FallbackBuffer Is Nothing Then
        Set FallbackBuffer = mDecoderFallback.CreateFallbackBuffer
    End If
    
    If FallbackBuffer.Fallback(ByteBuffer, 0) Then
        Dim Ch As Integer
        
        If CharIndex + FallbackBuffer.Remaining - 1 > UBound(Chars) Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowChars, EncodingName, "DecoderFallback"), "Chars")
        
        Do While FallbackBuffer.Remaining > 0
            Ch = FallbackBuffer.GetNextChar
            Chars(CharIndex) = Ch
            CharIndex = CharIndex + 1
        Loop
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Init .ReadProperty("BigEndian", False), .ReadProperty("EmitBOM", True), False
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BigEndian", mIsBigEndian
        .WriteProperty "EmitBOM", mEmitBOM
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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
'   Encoding Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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


