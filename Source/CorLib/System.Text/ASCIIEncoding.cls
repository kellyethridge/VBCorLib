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

Private Const ASCIIStart            As Long = 0
Private Const ASCIIEnd              As Long = 127
Private Const IsReadOnlyProp        As String = "IsReadOnly"
Private Const EncoderFallbackProp   As String = "EncoderFallback"
Private Const DecoderFallbackProp   As String = "DecoderFallback"

Private mIsReadOnly         As Boolean
Private mEncoderFallback    As EncoderFallback
Private mDecoderFallback    As DecoderFallback
Private mBytesUnknown(0)    As Byte


''
' When implemented in a derived class, gets a value indicating whether the current encoding is read-only.
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
Public Function Clone() As ASCIIEncoding
    Set Clone = New ASCIIEncoding
    Clone.InitClone Me
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
' @return Returns True for single-byte, False for multi-byte code points.
Public Property Get IsSingleByte() As Boolean
    IsSingleByte = True
End Property

''
' Gets the encoding name to be used in with the mail agent body tags.
'
' @return The body tag compatible encoding name.
' @see Encoding
Public Property Get BodyName() As String
    BodyName = Table.BodyName(CodePage)
End Property

''
' Gets the code page identifier for this encoding.
'
' @return Code page identifier.
' @see Encoding
Public Property Get CodePage() As Long
    CodePage = 20127
End Property

''
' Gets the name of this encoding.
'
' @return The encoding's name.
' @see Encoding
Public Property Get EncodingName() As String
    EncodingName = Table.EncodingName(CodePage)
End Property

''
' Gets the encoding name to be used in with the mail agent header tags.
'
' @return The header tag compatible encoding name.
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
' @return The registered encoding name.
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
' @param Chars A set of characters used to calculated the number of bytes once encoded with this encoding.
' @param Index The starting index of the first character to be used.
' @param Count The number of characters to include in the calculation.
' @return The number of bytes that would be produced by encoding the set of characters.
' @remarks The <i>Chars</i> parameter will accept either a String or an Integer array.
' <p>If a string is passed in, then the Index and Count parameters are ignored.</p>
' @see Encoding
Public Function GetByteCount(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim LocalChars()    As Integer
    Dim Range           As ListRange
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    Range = ValidateGetBytes(LocalChars, Index, Count)
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
' @see Encoding
Public Function GetBytes(ByRef Chars As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    Dim LocalChars()    As Integer
    Dim Range           As ListRange
    Dim ByteCount       As Long
    
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    
    Range = ValidateGetBytes(LocalChars, Index, Count)
    ByteCount = GetByteCountCore(LocalChars, Range.Index, Range.Count)
    SAPtr(GetBytes) = SafeArrayCreateVectorEx(vbByte, 0, ByteCount)
    GetBytesCore LocalChars, Range.Index, Range.Count, GetBytes, 0
    
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
' set of characters. To obtain the size required, use <a href="ASCIIEncoding.GetByteCount.html">GetByteCount</a> to find the necessary size.
' Or, use <a href="ASCIIEncoding.GetMaxByteCount.html">GetMaxByteCount</a> to get the largest size that could be produced. The largest size is not
' the perfect size, only the maximum number of bytes that could ever be produced by a number of
' characters. Resizing of the resultant <i>Bytes</i> array may be necessary.
' @see Encoding
Public Function GetBytesEx(ByRef Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Dim LocalChars() As Integer
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    ValidateGetBytesEx LocalChars, CharIndex, CharCount, Bytes, ByteIndex
    
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
' @see Encoding
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
' @see Encoding
Public Function GetChars(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Integer()
    Dim Range       As ListRange
    Dim CharCount   As Long
    
    Range = ValidateGetChars(Bytes, Index, Count)
    
    CharCount = GetCharCountCore(Bytes, Range.Index, Range.Count)
    SAPtr(GetChars) = SafeArrayCreateVectorEx(vbInteger, 0, CharCount)
    GetCharsCore Bytes, Range.Index, Range.Count, GetChars, 0
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
    GetCharsEx = GetCharsCore(Bytes, ByteIndex, ByteCount, Chars, CharIndex)
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
    Dim Result As Currency
    
    If ByteCount < 0 Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_NeedNonNegNum
    
    Result = ByteCount
    
    If mDecoderFallback.MaxCharCount > 1 Then
        Result = Result * mDecoderFallback.MaxCharCount
    End If
    
    If Result > &H7FFFFFFF Then _
        Error.ArgumentOutOfRange "ByteCount", ArgumentOutOfRange_GetCharCountOverflow
    
    GetMaxCharCount = Result
End Function

''
' Returns the maximum number of bytes that can be created from a specific number of characters.
'
' @param CharCount The number of characters to be encoded.
' @return The maximum number of bytes that can be generated from the specified number of characters.
' @see Encoding
Public Function GetMaxByteCount(ByVal CharCount As Long) As Long
    Dim Result As Currency
    
    If CharCount < 0 Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_NeedNonNegNum
    
    Result = CharCount + 1@
    
    If mEncoderFallback.MaxCharCount > 1 Then
        Result = Result * mEncoderFallback.MaxCharCount
    End If
    
    If Result > &H7FFFFFFF Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_GetByteCountOverflow
    
    GetMaxByteCount = Result
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
    Set GetEncoder = Cor.NewDefaultEncoder(Me)
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
    Set GetDecoder = Cor.NewDefaultDecoder(Me)
End Function

''
' Returns an array of bytes that represents this encoding.
'
' @return A byte array containing a sequence of bytes that specifies the encoding used.<br/>
' -or-<br/>
' A byte array of length zero, if a preamble is not required.
' <p>See <a href="Encoding.GetPreamble.html">Encoding.GetPreamble</a> for additional information.</p>
' @see Encoding
Public Function GetPreamble() As Byte()
    GetPreamble = Cor.NewBytes()
End Function

''
' Decodes a range of bytes into a String.
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
    Dim Range   As ListRange
    
    Range = ValidateGetChars(Bytes, Index, Count)
    GetString = SysAllocStringLen(0, Range.Count)
    Chars = AllocChars(GetString)
    
    On Error GoTo Catch
    GetCharsCore Bytes, Range.Index, Range.Count, Chars, 0
    
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
    Dim Other As ASCIIEncoding
    
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
Friend Sub InitClone(ByVal Other As ASCIIEncoding)
    Set mEncoderFallback = Other.EncoderFallback
    Set mDecoderFallback = Other.DecoderFallback
    mIsReadOnly = False
End Sub

Private Sub Class_InitProperties()
    Set mEncoderFallback = Statics.EncoderFallback.ReplacementFallback
    Set mDecoderFallback = Statics.DecoderFallback.ReplacementFallback
    mIsReadOnly = True
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mIsReadOnly = PropBag.ReadProperty(IsReadOnlyProp)
    Set mEncoderFallback = PropBag.ReadProperty(EncoderFallbackProp, Statics.EncoderFallback.ReplacementFallback)
    Set mDecoderFallback = PropBag.ReadProperty(DecoderFallbackProp, Statics.DecoderFallback.ReplacementFallback)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty IsReadOnlyProp, mIsReadOnly

    If Not mEncoderFallback Is Statics.EncoderFallback.ReplacementFallback Then
        PropBag.WriteProperty EncoderFallbackProp, mEncoderFallback
    End If
    If Not mDecoderFallback Is Statics.DecoderFallback.ReplacementFallback Then
        PropBag.WriteProperty DecoderFallbackProp, mDecoderFallback
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Table() As EncodingTable
    Set Table = Statics.Encoding.EncodingTable
End Property

Private Function GetByteCountCore(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    Dim i           As Long
    Dim Ch          As Integer
    Dim ByteCount   As Long
    Dim Fallback    As EncoderFallbackBuffer
    
    For i = Index To Index + Count - 1
        Ch = Chars(i)
        
        Select Case Ch
            ' we need to have a range because by using Integers to
            ' represent unicode characters, we may deal with negative values.
            Case ASCIIStart To ASCIIEnd
                ByteCount = ByteCount + 1
            Case Else
                ByteCount = ByteCount + GetEncoderFallbackCount(Ch, Fallback)
        End Select
    Next
    
    GetByteCountCore = ByteCount
End Function

Private Function GetEncoderFallbackCount(ByVal Ch As Integer, ByRef Fallback As EncoderFallbackBuffer) As Long
    If Fallback Is Nothing Then
        Set Fallback = mEncoderFallback.CreateFallbackBuffer
    End If
    
    If Fallback.Fallback(Ch) Then
        GetEncoderFallbackCount = Fallback.Remaining
        Fallback.Reset
    End If
End Function

Private Function GetBytesCore(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Dim Ch          As Long
    Dim CharEnd     As Long
    Dim ByteStart   As Long
    Dim ByteEnd     As Long
    Dim Fallback    As EncoderFallbackBuffer
    
    If Not TryGetBytesOptimized(Chars, CharIndex, CharCount, Bytes, ByteIndex, mEncoderFallback, GetBytesCore) Then
        ByteStart = ByteIndex
        ByteEnd = UBound(Bytes)
        CharEnd = CharIndex + CharCount
        
        Do While CharIndex < CharEnd
            Ch = Chars(CharIndex)
            
            Select Case Ch
                ' we need to have a range because by using Integers to
                ' represent unicode characters, we may deal with negative values.
                Case ASCIIStart To ASCIIEnd
                    If ByteIndex > ByteEnd Then _
                        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowBytes, EncodingName, TypeName(mEncoderFallback)), "Bytes")
                    
                    Bytes(ByteIndex) = Ch
                    ByteIndex = ByteIndex + 1
                Case Else
                    ByteIndex = InsertEncoderFallback(Ch, Bytes, ByteIndex, Fallback)
            End Select
            
            CharIndex = CharIndex + 1
        Loop
        
        GetBytesCore = ByteIndex - ByteStart
    End If
End Function

Private Function InsertEncoderFallback(ByVal Ch As Integer, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByRef Fallback As EncoderFallbackBuffer) As Long
    If Fallback Is Nothing Then
        Set Fallback = mEncoderFallback.CreateFallbackBuffer
    End If
    
    If Fallback.Fallback(Ch) Then
        If ByteIndex + Fallback.Remaining - 1 > UBound(Bytes) Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowBytes, EncodingName, TypeName(mEncoderFallback)), "Bytes")
        
        Do While Fallback.Remaining > 0
            Bytes(ByteIndex) = Fallback.GetNextChar
            ByteIndex = ByteIndex + 1
        Loop
    End If
    
    InsertEncoderFallback = ByteIndex
End Function

Private Function TryGetBytesOptimized(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal Fallback As EncoderFallback, ByRef ByteCount As Long) As Boolean
    Dim ReplacementChar As Integer
    Dim i               As Long
    Dim Ch              As Integer
    
    If TypeOf Fallback Is EncoderReplacementFallback And Fallback.MaxCharCount = 1 Then
        Dim EncoderFallback As EncoderReplacementFallback
        Set EncoderFallback = Fallback
        ReplacementChar = Asc(EncoderFallback.DefaultString)
        
        If ReplacementChar >= ASCIIStart And ReplacementChar <= ASCIIEnd Then
            If ByteIndex + CharCount - 1 > UBound(Bytes) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowBytes, EncodingName, TypeName(mEncoderFallback)), "Bytes")
            
            For i = CharIndex To CharIndex + CharCount - 1
                Ch = Chars(i)
                
                Select Case Ch
                    Case ASCIIStart To ASCIIEnd
                        Bytes(ByteIndex) = Ch
                    Case Else
                        Bytes(ByteIndex) = ReplacementChar
                End Select
                
                ByteIndex = ByteIndex + 1
            Next
            
            ByteCount = CharCount
            TryGetBytesOptimized = True
        End If
    End If
End Function

Private Function GetCharCountCore(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long) As Long
    Dim i           As Long
    Dim b           As Byte
    Dim CharCount   As Long
    Dim Fallback    As DecoderFallbackBuffer
    
    For i = ByteIndex To ByteIndex + ByteCount - 1
        b = Bytes(i)
        
        If b <= ASCIIEnd Then
            CharCount = CharCount + 1
        Else
            CharCount = CharCount + GetDecoderFallbackCount(b, Fallback)
        End If
    Next
    
    GetCharCountCore = CharCount
End Function

Private Function GetDecoderFallbackCount(ByVal b As Byte, ByRef Fallback As DecoderFallbackBuffer) As Long
    Dim Bytes(0) As Byte
    
    If Fallback Is Nothing Then
        Set Fallback = mDecoderFallback.CreateFallbackBuffer
    End If
    
    Bytes(0) = b
    If Fallback.Fallback(Bytes, 0) Then
        GetDecoderFallbackCount = Fallback.Remaining
        Fallback.Reset
    End If
End Function

Private Function GetCharsCore(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim b           As Byte
    Dim ByteEnd     As Long
    Dim Fallback    As DecoderFallbackBuffer
    
    ByteEnd = ByteIndex + ByteCount

    Do While ByteIndex < ByteEnd
        b = Bytes(ByteIndex)
        
        If b <= ASCIIEnd Then
            If CharIndex > UBound(Chars) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowChars, EncodingName, TypeName(mDecoderFallback)), "Chars")
        
            Chars(CharIndex) = b
            CharIndex = CharIndex + 1
        Else
            CharIndex = InsertDecoderChars(b, Chars, CharIndex, Fallback)
        End If
        
        ByteIndex = ByteIndex + 1
    Loop
    
    GetCharsCore = ByteCount
End Function

Private Function InsertDecoderChars(ByVal b As Byte, ByRef Chars() As Integer, ByVal CharIndex As Long, ByRef Fallback As DecoderFallbackBuffer) As Long
    If Fallback Is Nothing Then
        Set Fallback = mDecoderFallback.CreateFallbackBuffer
    End If
    
    mBytesUnknown(0) = b
    
    If Fallback.Fallback(mBytesUnknown, 0) Then
        If CharIndex + Fallback.Remaining - 1 > UBound(Chars) Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EncodingConversionOverflowChars, EncodingName, TypeName(mDecoderFallback)), "Chars")
        
        Do While Fallback.Remaining
            Chars(CharIndex) = Fallback.GetNextChar
            CharIndex = CharIndex + 1
        Loop
    End If
    
    InsertDecoderChars = CharIndex
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


