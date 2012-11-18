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
'    CopyRight (c) 2005 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: UnicodeEncoding
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

Private Const CHAR_SIZE As Long = 2

Private mChars()        As Integer
Private mCharsSA        As SafeArray1d
Private mIsBigEndian    As Boolean
Private mEmitBOM        As Boolean
Private mCodePage       As Long



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
    Encoding = "Unicode"
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
    
    GetByteCount = InternalGetByteCount(ElemCount)
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
    Ret = cArray.CreateInstance(ciByte, ElemCount * CHAR_SIZE)
    Call InternalGetBytes(ElemIndex, ElemCount, Ret, 0)
    
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
    
    Result = VerifyArrayRange(SAPtr(Bytes), ByteIndex, InternalGetByteCount(CharCount))
    Select Case Result
        Case Argument_InvalidOffLen:   Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Bytes")
        Case Is <> NO_ERROR:                ThrowArrayRangeException Result, "Bytes", ByteIndex, "ByteIndex", 0, vbNullString
    End Select
    
    GetBytesEx = InternalGetBytes(CharIndex, CharCount, Bytes, ByteIndex)
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
    Dim ElemCount As Long
    Dim ElemIndex As Long
    Dim Result As Long
    
    Result = GetOptionalArrayRange(SAPtr(Bytes), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Bytes", ElemIndex, "Index", ElemCount, "Count"
    
    GetCharCount = InternalGetCharCount(ElemCount)
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
    Ret = cArray.CreateInstance(ciInteger, ElemCount \ CHAR_SIZE)
    Call InternalGetChars(Bytes, ElemIndex, ElemCount, Ret, 0)
    
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

    Result = VerifyArrayRange(SAPtr(Chars), CharIndex, InternalGetCharCount(ByteCount))
    Select Case Result
        Case Argument_InvalidOffLen:   Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Chars")
        Case Is <> NO_ERROR:                ThrowArrayRangeException Result, "Chars", CharIndex, "CharIndex", 0, vbNullString
    End Select
    
    GetCharsEx = InternalGetChars(Bytes, ByteIndex, ByteCount, Chars, CharIndex)
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

    GetMaxByteCount = CharCount * CHAR_SIZE
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

    GetMaxCharCount = ByteCount \ CHAR_SIZE
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
    Dim ElemIndex As Long
    Dim ElemCount As Long
    Dim Result As Long
    
    Result = GetOptionalArrayRange(SAPtr(Bytes), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Bytes", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index)
    
    Dim size As Long
    size = InternalGetCharCount(ElemCount)
    GetString = SysAllocStringLen(0, size)
    
    With mCharsSA
        .pvData = StrPtr(GetString)
        .lLbound = 0
        .cElements = size
    End With
    SAPtr(mChars) = VarPtr(mCharsSA)
    Call InternalGetChars(Bytes, ElemIndex, ElemCount, mChars, 0)
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
    Dim Ret As New DefaultEncoder
    Call Ret.Init(Me)
    Set GetEncoder = Ret
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
    Dim Ret As New UnicodeDecoder
    Call Ret.Init(mIsBigEndian)
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
        
        If TypeOf Value Is UnicodeEncoding Then
            Dim en As UnicodeEncoding
            Set en = Value
            Equals = (en.GetHashCode = GetHashCode)
        End If
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = mCodePage Or CLng(IIf(mIsBigEndian, &H80000000, 0)) Or CLng(IIf(mEmitBOM, &H40000000, 0))
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal BigEndian As Boolean, ByVal ByteOrderMark As Boolean)
    mIsBigEndian = BigEndian
    mEmitBOM = ByteOrderMark
    If BigEndian Then mCodePage = 1201
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InternalGetBytes(ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    If CharCount = 0 Then Exit Function
    
    If mIsBigEndian Then
        Dim i As Long
        Dim ch As Long
        For i = 0 To CharCount - 1
            ch = mChars(CharIndex + i)
            Bytes(ByteIndex) = (ch \ &H100) And &HFF
            Bytes(ByteIndex + 1) = ch And &HFF
            ByteIndex = ByteIndex + 2
        Next i
    Else
        Call CopyMemory(Bytes(ByteIndex), mChars(CharIndex), CharCount * CHAR_SIZE)
    End If
    
    InternalGetBytes = CharCount * CHAR_SIZE
End Function

Private Function InternalGetChars(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    If ByteCount = 0 Then Exit Function
    
    If mIsBigEndian Then
        Dim i As Long
        Dim ch As Long
        For i = 0 To ByteCount - 1 Step 2
            ch = Bytes(ByteIndex + i) * &H100
            ch = ch Or Bytes(ByteIndex + i + 1)
            Chars(CharIndex) = AsWord(ch)
            CharIndex = CharIndex + 1
        Next i
    Else
        If ByteCount And 1 Then ByteCount = ByteCount - 1
        Call CopyMemory(Chars(CharIndex), Bytes(ByteIndex), ByteCount)
    End If
    InternalGetChars = InternalGetCharCount(ByteCount)
End Function

Private Function InternalGetByteCount(ByVal CharCount As Long) As Long
    InternalGetByteCount = CharCount * CHAR_SIZE
End Function

Private Function InternalGetCharCount(ByVal ByteCount As Long) As Long
    InternalGetCharCount = ByteCount \ CHAR_SIZE
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    mCodePage = 1200
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Call Init(.ReadProperty("BigEndian", False), .ReadProperty("EmitBOM", True))
    End With
End Sub

Private Sub Class_Terminate()
    SAPtr(mChars) = 0
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BigEndian", mIsBigEndian)
        Call .WriteProperty("EmitBOM", mEmitBOM)
    End With
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


