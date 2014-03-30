VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF7Encoding"
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
' Module: UTF7Encoding
'

''
' A set of functions used to convert unicode characters to and from byte arrays.
'
' @remarks Characters less that 128 may be directly encoded as their existing
' ascii value. Characters above 127 are encoded using a Base-64 encoding scheme.
' <p>Information about the UTF-7 encoding scheme can be found at "http://www.faqs.org/rfcs/rfc2152.html".</p>
'
' @see Constructors
' @see Encoding
'
Option Explicit
Implements IObject
Implements Encoding

Private Const CODE_PAGE As Long = 65000

Private mChars()                As Integer
Private mCharsSA                As SafeArray1d
Private mDirectlyEncodable()    As Boolean
Private mAllowOptionals         As Boolean



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
    BodyName = "utf-7"
End Property

''
' Returns the encoding name registered with the Internet Assigned Numbers Authority.
'
' @return The registered encoding name.
'
Public Property Get WebName() As String
    WebName = "utf-7"
End Property

''
' Returns the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
'
Public Property Get HeaderName() As String
    HeaderName = "utf-7"
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
    EncodingName = "Unicode (UTF-7)"
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
    IsBrowserSave = False
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
' Indicates if this encoding can be used to save date by mail and news clients.
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
    Call AttachChars(Chars, mChars, mCharsSA)
    
    Dim ElemIndex As Long
    Dim ElemCount As Long
    Dim Result As Long
    Result = GetOptionalArrayRange(SAPtr(mChars), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Chars", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index)
    
    GetByteCount = InternalGetByteCount(mChars, ElemIndex, ElemCount, Nothing)
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
    Call AttachChars(Chars, mChars, mCharsSA)
    
    Dim ElemIndex As Long
    Dim ElemCount As Long
    Dim Result As Long
    Result = GetOptionalArrayRange(SAPtr(mChars), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Chars", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index)
    
    Dim Ret() As Byte
    If ElemCount > 0 Then
        ReDim Ret(0 To InternalGetByteCount(mChars, ElemIndex, ElemCount, Nothing) - 1)
        Call InternalGetBytes(mChars, ElemIndex, ElemCount, Ret, 0, Nothing)
    Else
        Ret = Cor.NewBytes()
    End If
    
    GetBytes = Ret
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
    Call AttachChars(Chars, mChars, mCharsSA)
    
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(mChars), CharIndex, CharCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Chars", CharIndex, "CharIndex", CharCount, "CharCount"
    
    GetBytesEx = InternalGetBytes(mChars, CharIndex, CharCount, Bytes, ByteIndex, Nothing)
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
    Dim ElemIndex As Long
    Dim ElemCount As Long
    Dim Result As Long
    
    Result = GetOptionalArrayRange(SAPtr(Bytes), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Byte", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index)
    
    Dim Size As Long
    Size = GetMaxCharCount(ElemCount)
    
    Dim Ret As String
    Ret = SysAllocStringLen(0, Size)
    Call AttachChars(Ret, mChars, mCharsSA)
    Size = GetCharsEx(Bytes, ElemIndex, ElemCount, mChars, 0)
    GetString = Left$(Ret, Size)
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
    Dim ElemIndex As Long
    Dim ElemCount As Long
    Dim Result As Long
    
    Result = GetOptionalArrayRange(SAPtr(Bytes), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Bytes", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index)
    
    GetCharCount = InternalGetCharCount(Bytes, ElemIndex, ElemCount)
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
    Dim ElemIndex As Long
    Dim ElemCount As Long
    Dim Result As Long
    
    Result = GetOptionalArrayRange(SAPtr(Bytes), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Bytes", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index)
    
    Dim Ret() As Integer
    If ElemCount > 0 Then
        ReDim Ret(0 To GetMaxCharCount(ElemCount))
        
        Dim Size As Long
        Size = InternalGetChars(Bytes, ElemIndex, ElemCount, Ret, 0, Nothing)
        If Size <> UBound(Ret) + 1 Then ReDim Preserve Ret(0 To Size - 1)
    Else
        Ret = Cor.NewIntegers()
    End If
    
    GetChars = Ret
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
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(Bytes), ByteIndex, ByteCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Bytes", ByteIndex, "ByteIndex", ByteCount, "ByteCount"
    
    GetCharsEx = InternalGetChars(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Nothing)
End Function

''
' Returns the maximum number of characters than can be decoded from the number of bytes specified.
'
' @param ByteCount The number of bytes to be decoded.
' @return The maximum number of characters that can be decoded from the specified number of bytes.
'
Public Function GetMaxCharCount(ByVal ByteCount As Long) As Long
    If ByteCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "ByteCount", ByteCount)
        
    GetMaxCharCount = ByteCount
End Function

''
' Returns the maximum number of bytes that can be created from a specific number of characters.
'
' @param CharCount The number of characters to be encoded.
' @return The maximum number of bytes that can be generated from the specified number of characters.
'
Public Function GetMaxByteCount(ByVal CharCount As Long) As Long
    If CharCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "CharCount", CharCount)
    
    GetMaxByteCount = 2 + (-(Int(-((CharCount * 16) / 6))))
End Function

''
' Returns an array of bytes that represents this encoding.
'
' @return A byte array containg the bytes to be used to identify this encoding type.
' @remarks <p>There is no identifier for UTF7 so an empty array is returned.</p>
'
Public Function GetPreamble() As Byte()
    GetPreamble = Cor.NewBytes()
End Function

''
' Returns an encoder that maintains state.
'
' @return The stateful encoder.
' @remarks Unlike UTF7Encoding, the encoder maintains state between encoding calls. This allows
' for a set of characters to be split up but still be encoded as a single set of characters. This
' will allow Modified Base64 encoding to continue between encoding calls.
'
Public Function GetEncoder() As Encoder
    Dim Ret As New UTF7Encoder
    Call Ret.Init(Me)
    Set GetEncoder = Ret
End Function

''
' Returns a decoder that maintains state.
'
' @return The stateful decoder.
' @remarks Unlike UTF8Encoding, the decoder maintains state between decoding calls. this allows
' for a sequence of bytes to not have to be in the same call in order to be decoded to the
' representive character. All characters other than directly encodable characters are encoded
' in a Modified Base64 encoding, which can be require 3 bytes to decode to a single character.
' This decoder allows the Base64 encoded byte set to be decoded across multiple calls and still
' be decoded as a single set of bytes. A multi-byte encoded characters does not need to be in
' the same call in order to be decoded properly.
'
Public Function GetDecoder() As Decoder
    Dim Ret As New UTF7Decoder
    Call Ret.Init(Me)
    Set GetDecoder = Ret
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
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
        
        If TypeOf Value Is UTF7Encoding Then
            Dim en As UTF7Encoding
            Set en = Value
            Equals = (en.AllowOptionals = mAllowOptionals)
        End If
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = CODE_PAGE Or CLng(IIf(mAllowOptionals, &H80000000, 0))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal AllowOptionals As Boolean)
    Dim i As Long
    
    If AllowOptionals Then
        For i = 1 To 20
            mDirectlyEncodable(Choose(i, 33, 34, 35, 36, 37, 38, 42, 59, 60, 61, 62, 64, 91, 93, 94, 95, 96, 123, 124, 125)) = True
        Next i
    End If
    mAllowOptionals = AllowOptionals
End Sub

Friend Property Get AllowOptionals() As Boolean
    AllowOptionals = mAllowOptionals
End Property

Friend Function InternalGetBytes(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByRef Encoder As UTF7Encoder) As Long
    Dim Bits As Long
    Dim bitCount As Long
    Dim maxCharIndex As Long
    Dim MaxByteIndex As Long
    Dim ch As Long
    Dim ByteStartIndex As Long
    Dim directlyEncodable As Boolean
    
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException("Bytes", Environment.GetResourceString(ArgumentNull_Array))
    If ByteIndex < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "ByteIndex", ByteIndex)
    
    If Not Encoder Is Nothing Then
        With Encoder
            Bits = .Bits
            bitCount = .bitCount
        End With
    Else
        bitCount = -1
    End If
    
    ByteStartIndex = ByteIndex
    maxCharIndex = Index + Count - 1
    MaxByteIndex = UBound(Bytes)
    Do While Index <= maxCharIndex
        AsWord(ch) = Chars(Index)
        directlyEncodable = False
        If ch < &H80 Then
            If mDirectlyEncodable(ch) Then
                If bitCount >= 0 Then
                    If bitCount > 0 Then
                        If ByteIndex > MaxByteIndex Then Call SmallBufferError("Bytes")
                        Bytes(ByteIndex) = Base64Bytes(Helper.ShiftLeft(Bits, 6 - bitCount) And &H3F)
                        ByteIndex = ByteIndex + 1
                    End If
                    If ByteIndex > MaxByteIndex Then Call SmallBufferError("Bytes")
                    Bytes(ByteIndex) = vbMinus
                    ByteIndex = ByteIndex + 1
                End If
                If ByteIndex > MaxByteIndex Then Call SmallBufferError("Bytes")
                Bytes(ByteIndex) = ch
                ByteIndex = ByteIndex + 1
                bitCount = -1
                directlyEncodable = True
            End If
        End If
        If Not directlyEncodable Then
            If ch = vbPlus And bitCount < 0 Then
                If ByteIndex > MaxByteIndex Then Call SmallBufferError("Bytes")
                Bytes(ByteIndex) = vbPlus
                Bytes(ByteIndex + 1) = vbMinus
                ByteIndex = ByteIndex + 2
            Else
                If bitCount < 0 Then
                    If ByteIndex > MaxByteIndex Then Call SmallBufferError("Bytes")
                    Bytes(ByteIndex) = vbPlus
                    ByteIndex = ByteIndex + 1
                    bitCount = 0
                End If
                Bits = (Bits * &H10000) Or ch
                bitCount = bitCount + 16
                Do While bitCount >= 6
                    bitCount = bitCount - 6
                    If ByteIndex > MaxByteIndex Then Call SmallBufferError("Bytes")
                    Bytes(ByteIndex) = Base64Bytes(Helper.ShiftRight(Bits, bitCount) And &H3F)
                    ByteIndex = ByteIndex + 1
                Loop
                Bits = Bits And (Powers(bitCount) - 1)
            End If
        End If
        Index = Index + 1
    Loop
    
    Dim Flush As Boolean
    If Encoder Is Nothing Then
        Flush = True
    Else
        Flush = Encoder.Flush
    End If
    
    If Flush And (bitCount >= 0) Then
        If bitCount > 0 Then
            If ByteIndex > MaxByteIndex Then _
                Throw Cor.NewArgumentException("Bytes conversion buffer is too small.", "Bytes")
            
            Bytes(ByteIndex) = Base64Bytes(Helper.ShiftLeft(Bits, 6 - bitCount) And &H3F)
            ByteIndex = ByteIndex + 1
        End If
        If ByteIndex > MaxByteIndex Then _
            Throw Cor.NewArgumentException("Bytes conversion buffer is too small.", "Bytes")
        
        Bytes(ByteIndex) = vbMinus
        ByteIndex = ByteIndex + 1
        bitCount = -1
    End If
    
    If Not Encoder Is Nothing Then
        With Encoder
            .Bits = Bits
            .bitCount = bitCount
        End With
    End If
    
    InternalGetBytes = ByteIndex - ByteStartIndex
End Function

Friend Function InternalGetByteCount(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByRef Encoder As UTF7Encoder) As Long
    Dim bitCount As Long
    Dim maxCharIndex As Long
    Dim ch As Long
    Dim directlyEncodable As Boolean
    Dim ByteIndex As Long
    
    If Not Encoder Is Nothing Then
        bitCount = Encoder.bitCount
    Else
        bitCount = -1
    End If
    
    maxCharIndex = Index + Count - 1
    Do While Index <= maxCharIndex
        AsWord(ch) = Chars(Index)
        directlyEncodable = False
        If ch < &H80 Then
            If mDirectlyEncodable(ch) Then
                If bitCount >= 0 Then
                    If bitCount > 0 Then ByteIndex = ByteIndex + 1
                    ByteIndex = ByteIndex + 1
                End If
                ByteIndex = ByteIndex + 1
                bitCount = -1
                directlyEncodable = True
            End If
        End If
        If Not directlyEncodable Then
            If ch = vbPlus And bitCount < 0 Then
                ByteIndex = ByteIndex + 2
            Else
                If bitCount < 0 Then
                    ByteIndex = ByteIndex + 1
                    bitCount = 0
                End If
                bitCount = bitCount + 16
                Do While bitCount >= 6
                    bitCount = bitCount - 6
                    ByteIndex = ByteIndex + 1
                Loop
            End If
        End If
        Index = Index + 1
    Loop
    
    Dim Flush As Boolean
    If Encoder Is Nothing Then
        Flush = True
    Else
        Flush = Encoder.Flush
    End If
    
    If Flush And (bitCount >= 0) Then
        If bitCount > 0 Then
            ByteIndex = ByteIndex + 1
        End If
        ByteIndex = ByteIndex + 1
        bitCount = -1
    End If
    
    InternalGetByteCount = ByteIndex
End Function

Friend Function InternalGetCharCount(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Long
    Dim InBase64 As Boolean
    Dim MaxIndex As Long
    Dim CharCount As Long
    Dim b As Byte
    Dim Bits As Currency
    
    MaxIndex = Index + Count
    Do While Index < MaxIndex
        b = Bytes(Index)
        If Not InBase64 Then
            If b = vbPlus Then
                InBase64 = True
            Else
                CharCount = CharCount + 1
            End If
        ElseIf b = vbMinus Then
            If Bytes(Index - 1) = vbPlus Then CharCount = CharCount + 1
            CharCount = CharCount + Int(Bits / 16@)
            Bits = 0@
            InBase64 = False
        Else
            Bits = Bits + 6@
        End If
        Index = Index + 1
    Loop
    
    InternalGetCharCount = Int(Bits / 16@) + CharCount
End Function

Friend Function InternalGetChars(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByRef Decoder As UTF7Decoder) As Long
    Dim Bits As Long
    Dim bitCount As Long
    Dim b As Byte
    Dim InBase64 As Boolean
    Dim HasPlus As Boolean
    Dim CharStartIndex As Long
    Dim maxCharIndex As Long
    Dim MaxByteIndex As Long
    
    If cArray.IsNull(Chars) Then _
        Throw Cor.NewArgumentNullException("Chars", Environment.GetResourceString(ArgumentNull_Array))
    If CharIndex < LBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "CharIndex", CharIndex)
    
    If Not Decoder Is Nothing Then
        With Decoder
            Bits = .Bits
            bitCount = .bitCount
            InBase64 = .InBase64
            HasPlus = .HasPlus
        End With
    End If
    
    CharStartIndex = CharIndex
    maxCharIndex = UBound(Chars)
    MaxByteIndex = ByteIndex + ByteCount - 1
    Do While ByteIndex <= MaxByteIndex
        b = Bytes(ByteIndex)
        'Debug.Assert b <> 44
        Select Case b
            Case vbMinus
                If HasPlus Then
                    If CharIndex > maxCharIndex Then Call SmallBufferError("Chars")
                    Chars(CharIndex) = vbPlus
                    CharIndex = CharIndex + 1
                ElseIf InBase64 Then
                    bitCount = 0
                    Bits = 0
                    InBase64 = False
                Else
                    If CharIndex > maxCharIndex Then Call SmallBufferError("Chars")
                    Chars(CharIndex) = vbMinus
                    CharIndex = CharIndex + 1
                End If
                HasPlus = False
            Case vbPlus
                If InBase64 Then
                    Bits = (Bits * &H40) Or Base64CharToBits(b)
                    bitCount = bitCount + 6
                    HasPlus = False
                Else
                    Bits = 0
                    bitCount = 0
                    InBase64 = True
                    HasPlus = True
                End If
            Case Is < &H80
                If InBase64 Then
                    If Base64CharToBits(b) >= 0 Then
                        Bits = (Bits * &H40) Or Base64CharToBits(b)
                        bitCount = bitCount + 6
                    Else
                        If CharIndex > maxCharIndex Then Call SmallBufferError("Chars")
                        Chars(CharIndex) = b
                        CharIndex = CharIndex + 1
                        bitCount = 0
                        Bits = 0
                        InBase64 = False
                    End If
                Else
                    If CharIndex > maxCharIndex Then Call SmallBufferError("Chars")
                    Chars(CharIndex) = b
                    CharIndex = CharIndex + 1
                End If
                HasPlus = False
            Case Else
                If CharIndex > maxCharIndex Then Call SmallBufferError("Chars")
                Chars(CharIndex) = b
                CharIndex = CharIndex + 1
                bitCount = 0
                Bits = 0
                InBase64 = False
                HasPlus = False
        End Select
        If bitCount >= 16 Then
            If CharIndex > maxCharIndex Then Call SmallBufferError("Chars")
            Chars(CharIndex) = AsWord(Helper.ShiftRight(Bits, bitCount - 16) And &HFFFF&)
            CharIndex = CharIndex + 1
            bitCount = bitCount - 16
        End If
        ByteIndex = ByteIndex + 1
        Bits = Bits And &HFFFFFF
    Loop
    
    If Not Decoder Is Nothing Then
        With Decoder
            .bitCount = bitCount
            .Bits = Bits
            .HasPlus = HasPlus
            .InBase64 = InBase64
        End With
    End If
    
    InternalGetChars = CharIndex - CharStartIndex
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SmallBufferError(ByVal ParamName As String)
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), ParamName)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim i As Long
    
    ReDim mDirectlyEncodable(0 To 127)
    For i = 0 To 25
        mDirectlyEncodable(65 + i) = True
        mDirectlyEncodable(97 + i) = True
    Next i
    For i = 0 To 9
        mDirectlyEncodable(48 + i) = True
    Next i
    For i = 1 To 13
        mDirectlyEncodable(Choose(i, 9, 10, 13, 32, 39, 40, 41, 44, 45, 46, 47, 58, 63)) = True
    Next i

    With mCharsSA
        .cbElements = 2
        .cDims = 1
    End With
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Call Init(PropBag.ReadProperty("AllowOptionals", False))
End Sub

Private Sub Class_Terminate()
    SAPtr(mChars) = 0
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AllowOptionals", mAllowOptionals)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Encoding Interface
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


