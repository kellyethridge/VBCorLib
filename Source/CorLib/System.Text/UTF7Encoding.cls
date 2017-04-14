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

Private Const EncodableCharacters   As String = "\t\n\r '(),-./0123456789:?ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Private Const OptionalCharacters    As String = "!""#$%&*;<=>@[]^_`{|}"

Private mIsReadOnly                     As Boolean
Private mDirectlyEncodable(0 To 127)    As Boolean
Private mAllowOptionals                 As Boolean
Private mDecoderFallback                As DecoderFallback
Private mEncoderFallback                As EncoderFallback



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
' @return A cloned instance of a <a href="UTF7Encoding.html">UTF7Encoding</a> object.
' @remarks A cloned instance will have the read-only limitations removed.
'
Public Function Clone() As UTF7Encoding
    Set Clone = New UTF7Encoding
    Clone.InitClone mAllowOptionals
End Function

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
    CodePage = 65000
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
' Indicates if this encoding can be used to save date by mail and news clients.
'
' @return Indication for the mail and news clients to use this encoding for saving.
'
Public Property Get IsMailNewsSave() As Boolean
    IsMailNewsSave = Table.IsMailNewsSave(CodePage)
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
    
    On Error GoTo Catch
    
    LocalChars = AsChars(Chars)
    Range = ValidateGetBytes(LocalChars, Index, Count)
    
    If Range.Count > 0 Then
        GetByteCount = GetByteCountCore(LocalChars, Range.Index, Range.Count, Nothing)
    End If
    
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
    
    On Error GoTo Catch
    
    LocalChars = AsChars(Chars)
    Range = ValidateGetBytes(LocalChars, Index, Count)
    
    If Range.Count > 0 Then
        ReDim Output(0 To GetByteCountCore(LocalChars, Range.Index, Range.Count, Nothing) - 1)
        GetBytesCore LocalChars, Range.Index, Range.Count, Output, 0, Nothing
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
' Decodes a set of bytes into a String.
'
' @param Bytes The set of bytes to be decoded into a string.
' @param Index The index of the first byte to be decoded.
' @param Count The number of bytes to be used in the decoding.
' @return A string containing the decoded set of bytes.
'
Public Function GetString(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As String
    Dim Range   As ListRange
    Dim Output() As Integer
    
    Range = ValidateGetChars(Bytes, Index, Count)
    GetString = SysAllocStringLen(0, GetCharCountCore(Bytes, Range.Index, Range.Count))
    Output = AllocChars(GetString)
    
    On Error GoTo Catch
    GetCharsCore Bytes, Range.Index, Range.Count, Output, 0, Nothing
    
    GoSub Finally
    Exit Function

Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    FreeChars Output
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
    
    GetCharCount = GetCharCountCore(Bytes, Range.Index, Range.Count)
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
    Dim Output()    As Integer
    
    Range = ValidateGetChars(Bytes, Index, Count)
    CharCount = GetCharCountCore(Bytes, Range.Index, Range.Count)
    
    If CharCount > 0 Then
        ReDim Output(0 To CharCount - 1)
        GetCharsCore Bytes, Range.Index, Range.Count, Output, 0, Nothing
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
    ValidateGetCharsEx Bytes, ByteIndex, ByteCount, Chars, CharIndex
    GetCharsEx = GetCharsCore(Bytes, ByteIndex, ByteCount, Chars, CharIndex, Nothing)
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
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_NeedNonNegNum
    
    GetMaxByteCount = CharCount * 3 + 2
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
    Dim Encoder As New UTF7Encoder
    Encoder.Init Me
    Set GetEncoder = Encoder
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
    Dim Decoder As New UTF7Decoder
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
    Dim Other As UTF7Encoding
    
    On Error GoTo Catch
    Set Other = Value
    Equals = (Other.AllowOptionals = mAllowOptionals)
    
Catch:
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = CodePage Or IIfLong(mAllowOptionals, &H80000000, 0)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal AllowOptionals As Boolean)
    If AllowOptionals Then
        InitEncodable OptionalCharacters
        mAllowOptionals = True
    End If
End Sub

Friend Sub InitClone(ByVal AllowOptionals As Boolean)
    Init AllowOptionals
    mIsReadOnly = False
End Sub

Private Sub Class_Initialize()
    InitEncodable Funcs.t(EncodableCharacters)
    Set mEncoderFallback = Cor.NewEncoderReplacementFallback("")
    Set mDecoderFallback = New UTF7DecoderFallback
    mIsReadOnly = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitEncodable(ByRef s As String)
    Dim i       As Long
    Dim Chars() As Integer
    
    Chars = AllocChars(s)
    
    For i = 0 To UBound(Chars)
        mDirectlyEncodable(Chars(i)) = True
    Next
    
    FreeChars Chars
End Sub

Private Property Get Table() As EncodingTable
    Set Table = Statics.Encoding.EncodingTable
End Property

Friend Property Get AllowOptionals() As Boolean
    AllowOptionals = mAllowOptionals
End Property

Friend Function GetBytesCore(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal Encoder As UTF7Encoder) As Long
    Dim Bits                As Long
    Dim BitCount            As Long
    Dim MaxCharIndex        As Long
    Dim MaxByteIndex        As Long
    Dim Ch                  As Long
    Dim ByteStartIndex      As Long
    Dim DirectlyEncodable   As Boolean
        
    If Not Encoder Is Nothing Then
        With Encoder
            Bits = .Bits
            BitCount = .BitCount
        End With
    Else
        BitCount = -1
    End If
    
    ByteStartIndex = ByteIndex
    MaxCharIndex = Index + Count - 1
    MaxByteIndex = UBound(Bytes)
    
    Do While Index <= MaxCharIndex
        Ch = Chars(Index) And &HFFFF&
        DirectlyEncodable = False
        
        If Ch < &H80 Then
            If mDirectlyEncodable(Ch) Then
                If BitCount >= 0 Then
                    If BitCount > 0 Then
                        If ByteIndex > MaxByteIndex Then _
                            ThrowEncoderSmallConversionBuffer
                        
                        Bytes(ByteIndex) = Base64Bytes((Bits * Powers(6 - BitCount)) And &H3F)
                        ByteIndex = ByteIndex + 1
                    End If
                    
                    If ByteIndex > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer
                    
                    Bytes(ByteIndex) = vbMinusChar
                    ByteIndex = ByteIndex + 1
                End If
                
                If ByteIndex > MaxByteIndex Then _
                    ThrowEncoderSmallConversionBuffer
                
                Bytes(ByteIndex) = Ch
                ByteIndex = ByteIndex + 1
                BitCount = -1
                DirectlyEncodable = True
            End If
        End If
        
        If Not DirectlyEncodable Then
            If Ch = vbPlusChar And BitCount < 0 Then
                If ByteIndex > MaxByteIndex Then _
                    ThrowEncoderSmallConversionBuffer
                
                Bytes(ByteIndex) = vbPlusChar
                Bytes(ByteIndex + 1) = vbMinusChar
                ByteIndex = ByteIndex + 2
            Else
                If BitCount < 0 Then
                    If ByteIndex > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer
                    
                    Bytes(ByteIndex) = vbPlusChar
                    ByteIndex = ByteIndex + 1
                    BitCount = 0
                End If
                
                Bits = (Bits * vbShift16Bits) Or Ch
                BitCount = BitCount + 16
                
                Do While BitCount >= 6
                    BitCount = BitCount - 6
                    
                    If ByteIndex > MaxByteIndex Then _
                        ThrowEncoderSmallConversionBuffer
                    
                    Bytes(ByteIndex) = Base64Bytes((Bits \ (Powers(BitCount))) And &H3F)
                    ByteIndex = ByteIndex + 1
                Loop
                
                Bits = Bits And (Powers(BitCount) - 1)
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
    
    If Flush And (BitCount >= 0) Then
        If BitCount > 0 Then
            If ByteIndex > MaxByteIndex Then _
                ThrowEncoderSmallConversionBuffer
            
            Bytes(ByteIndex) = Base64Bytes((Bits * Powers(6 - BitCount)) And &H3F)
            ByteIndex = ByteIndex + 1
        End If
        
        If ByteIndex > MaxByteIndex Then _
            ThrowEncoderSmallConversionBuffer
        
        Bytes(ByteIndex) = vbMinusChar
        ByteIndex = ByteIndex + 1
        BitCount = -1
    End If
    
    If Not Encoder Is Nothing Then
        With Encoder
            .Bits = Bits
            .BitCount = BitCount
        End With
    End If
    
    GetBytesCore = ByteIndex - ByteStartIndex
End Function

Friend Function GetByteCountCore(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByVal Encoder As UTF7Encoder) As Long
    Dim BitCount            As Long
    Dim MaxCharIndex        As Long
    Dim Ch                  As Long
    Dim DirectlyEncodable   As Boolean
    Dim ByteCount           As Long
    
    If Not Encoder Is Nothing Then
        BitCount = Encoder.BitCount
    Else
        BitCount = -1
    End If
    
    MaxCharIndex = Index + Count - 1
    Do While Index <= MaxCharIndex
        Ch = Chars(Index) And &HFFFF&
        DirectlyEncodable = False
        
        If Ch < &H80 Then
            If mDirectlyEncodable(Ch) Then
                If BitCount > 0 Then
                    ByteCount = ByteCount + 3
                ElseIf BitCount = 0 Then
                    ByteCount = ByteCount + 2
                Else
                    ByteCount = ByteCount + 1
                End If
                
                BitCount = -1
                DirectlyEncodable = True
            End If
        End If
        
        If Not DirectlyEncodable Then
            If Ch = vbPlusChar And BitCount < 0 Then
                ByteCount = ByteCount + 2 ' 2 bytes for plus/minus combo
            Else
                If BitCount < 0 Then
                    ByteCount = ByteCount + 1
                    BitCount = 0
                End If
                
                BitCount = BitCount + 16
                
                Dim t As Long
                t = BitCount \ 6
                ByteCount = ByteCount + t
                BitCount = BitCount - t * 6
            End If
        End If
        
        Index = Index + 1
    Loop
    
    Dim Flush As Boolean
    If Encoder Is Nothing Then
        Flush = True
    ElseIf Encoder.Flush Then
        Flush = Encoder.Flush
    End If
    
    If Flush And (BitCount >= 0) Then
        If BitCount > 0 Then
            ByteCount = ByteCount + 1
        End If

        ByteCount = ByteCount + 1
    End If
    
    GetByteCountCore = ByteCount
End Function

Friend Function GetCharCountCore(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Long
    Dim InBase64    As Boolean
    Dim MaxIndex    As Long
    Dim CharCount   As Long
    Dim b           As Byte
    Dim BitCount    As Currency
    Dim Fallback    As DecoderFallbackBuffer
    
    MaxIndex = Index + Count
    Do While Index < MaxIndex
        b = Bytes(Index)
        
        If Not InBase64 Then
            If b = vbPlusChar Then
                InBase64 = True
            Else
                CharCount = CharCount + 1
            End If
        ElseIf b = vbMinusChar Then
            If Bytes(Index - 1) = vbPlusChar Then
                CharCount = CharCount + 1
            End If
            
            CharCount = CharCount + BitCount \ 16
            BitCount = 0
            InBase64 = False
        ElseIf b < &H80 Then
            ' check not valid base-64 character
            If IsBase64Byte(b) Then
                BitCount = BitCount + 6
            Else
                ' add fallback count
                CharCount = CharCount + GetFallbackCharCount(b, Fallback)
                BitCount = 0
                InBase64 = False
            End If
        Else
            ' add fallback count
            CharCount = CharCount + GetFallbackCharCount(b, Fallback)
            BitCount = 0
            InBase64 = False
        End If
        
        Index = Index + 1
    Loop
    
    GetCharCountCore = CharCount + BitCount \ 16
End Function

Private Function GetFallbackCharCount(ByVal b As Byte, ByRef Fallback As DecoderFallbackBuffer) As Long
    Dim BytesUnknown(0 To 0) As Byte
    
    If Fallback Is Nothing Then
        Set Fallback = mDecoderFallback.CreateFallbackBuffer
    End If
    
    BytesUnknown(0) = b
    If Fallback.Fallback(BytesUnknown, 0) Then
        GetFallbackCharCount = Fallback.Remaining
        Fallback.Reset
    End If
End Function

Friend Function GetCharsCore(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal Decoder As UTF7Decoder) As Long
    Dim Bits            As Long
    Dim BitCount        As Long
    Dim b               As Byte
    Dim InBase64        As Boolean
    Dim HasPlus         As Boolean
    Dim CharStartIndex  As Long
    Dim MaxCharIndex    As Long
    Dim MaxByteIndex    As Long
    Dim Fallback        As DecoderFallbackBuffer
    
    If Not Decoder Is Nothing Then
        With Decoder
            Bits = .Bits
            BitCount = .BitCount
            InBase64 = .InBase64
            HasPlus = .HasPlus
        End With
    End If
    
    CharStartIndex = CharIndex
    MaxCharIndex = UBound(Chars)
    MaxByteIndex = ByteIndex + ByteCount - 1
    
    Do While ByteIndex <= MaxByteIndex
        b = Bytes(ByteIndex)
        
        Select Case b
            Case vbMinusChar
                If HasPlus Then
                    If CharIndex > MaxCharIndex Then _
                        ThrowDecoderSmallConversionBuffer
                    
                    Chars(CharIndex) = vbPlusChar
                    CharIndex = CharIndex + 1
                ElseIf InBase64 Then
                    BitCount = 0
                    Bits = 0
                    InBase64 = False
                Else
                    If CharIndex > MaxCharIndex Then _
                        ThrowDecoderSmallConversionBuffer
                    
                    Chars(CharIndex) = vbMinusChar
                    CharIndex = CharIndex + 1
                End If
                
                HasPlus = False
            Case vbPlusChar
                If InBase64 Then
                    Bits = (Bits * vbShift6Bits) Or Base64CharToBits(b)
                    BitCount = BitCount + 6
                    HasPlus = False
                Else
                    Bits = 0
                    BitCount = 0
                    InBase64 = True
                    HasPlus = True
                End If
            Case Is < &H80
                If InBase64 Then
                    If Base64CharToBits(b) >= 0 Then
                        Bits = (Bits * vbShift6Bits) Or Base64CharToBits(b)
                        BitCount = BitCount + 6
                    Else
                        If CharIndex > MaxCharIndex Then _
                            ThrowDecoderSmallConversionBuffer
                        
                        Chars(CharIndex) = b
                        CharIndex = CharIndex + 1
                        BitCount = 0
                        Bits = 0
                        InBase64 = False
                    End If
                Else
                    If CharIndex > MaxCharIndex Then _
                        ThrowDecoderSmallConversionBuffer
                    
                    Chars(CharIndex) = b
                    CharIndex = CharIndex + 1
                End If
                
                HasPlus = False
            Case Else
                CharIndex = InsertFallbackChars(b, Chars, CharIndex, Fallback)
                BitCount = 0
                Bits = 0
                InBase64 = False
                HasPlus = False
        End Select
        
        If BitCount >= 16 Then
            If CharIndex > MaxCharIndex Then _
                ThrowDecoderSmallConversionBuffer
            
            Chars(CharIndex) = AsWord((Bits \ Powers(BitCount - 16)) And &HFFFF&)
            CharIndex = CharIndex + 1
            BitCount = BitCount - 16
        End If
        
        ByteIndex = ByteIndex + 1
        Bits = Bits And &HFFFFFF
    Loop
    
    If Not Decoder Is Nothing Then
        Decoder.BitCount = BitCount
        Decoder.Bits = Bits
        Decoder.HasPlus = HasPlus
        Decoder.InBase64 = InBase64
    End If
    
    GetCharsCore = CharIndex - CharStartIndex
End Function

Private Function InsertFallbackChars(ByVal b As Byte, ByRef Chars() As Integer, ByVal CharIndex As Long, ByRef Fallback As DecoderFallbackBuffer) As Long
    Dim BytesUnknown(0) As Byte
    
    If Fallback Is Nothing Then
        Set Fallback = mDecoderFallback.CreateFallbackBuffer
    End If
    
    BytesUnknown(0) = b
    If Fallback.Fallback(BytesUnknown, 0) Then
        Do While Fallback.Remaining > 0
            If CharIndex > UBound(Chars) Then _
                ThrowDecoderSmallConversionBuffer
            
            Chars(CharIndex) = Fallback.GetNextChar
            CharIndex = CharIndex + 1
        Loop
    End If
    
    InsertFallbackChars = CharIndex
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    Init PropBag.ReadProperty("AllowOptionals", False)
    Set mEncoderFallback = PropBag.ReadProperty("EncoderFallback")
    Set mDecoderFallback = PropBag.ReadProperty("DecoderFallback", New UTF7DecoderFallback)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AllowOptionals", mAllowOptionals
    PropBag.WriteProperty "EncoderFallback", mEncoderFallback
    
    If Not TypeOf mDecoderFallback Is UTF7DecoderFallback Then
        PropBag.WriteProperty "DecoderFallback", mDecoderFallback
    End If
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


