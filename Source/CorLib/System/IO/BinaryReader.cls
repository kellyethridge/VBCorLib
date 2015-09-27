VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: BinaryReader
'

''
' Provides a set of methods to read typed information from a Stream.
'
' @remarks
' <p>The bytes from a stream can be read and cast to a specific datatype. Each
' datatype will determine the number of bytes to be read from the stream. Once
' the number of bytes have been read, they are assembled into the requested
' datatype, advancing the position in the stream the required number of bytes.</p>
' <p>Characters can be read from the Stream using the provided Encoding system.
' The default encoding system is equal to New UTF8Encoding.</p>
' <p>To create a new <b>BinaryReader</b> use the public constructor.
' <pre>
' Set reader = Cor.NewBinaryReader(MyFileStream)
'
' - or -
'
' Set reader = NewBinaryReader(MyFileStream)
' </pre></p>
' @see Constructors
' @see BinaryWriter
'
Option Explicit
Implements IObject

Private Const DefaultCapacity   As Long = 16
Private Const MinDateAsDouble   As Double = -657434#
Private Const MaxDateAsDouble   As Double = 2958465.99998843
Private Const NoChars           As Long = -1

Private mStream         As Stream
Private mEncoding       As Encoding
Private mBuffer()       As Byte
Private mSingleChar(0)  As Integer


''
' Returns the underlying Stream being read by the Reader.
'
' @return The underlying Stream.
'
Public Property Get BaseStream() As Stream
    Set BaseStream = mStream
End Property

''
' Closes the reader and underlying Stream.
'
Public Sub CloseReader()
    If Not mStream Is Nothing Then
        mStream.CloseStream
        Set mStream = Nothing
    End If
End Sub

''
' Fills either a Byte array or Integer array with the specified number of elements. Or returns
' the next character to be decoded from the stream.
'
' @param Buffer The array to write the data to.
' @param Index The first index in <i>Buffer</i> to start writing data.
' @param Count The number of elements to be written to the array.
' @return The actual number of elements written to the array. This could be less than the requested amount.
' @remarks If a Byte array is passed in, then the bytes from the underlying stream are
' copied to the array. If an Integer array is passed in, then enough bytes are read from
' the stream to produce the requested number of decoded characters. The decoded characters
' are placed in the array starting at <i>Index</i>.
' <p>If all parameters are missing then a single character will be decoded from the stream and returned.</p>
' <p>All parameters must either be missing or supplied or an exception is thrown.</p>
'
Public Function Read(Optional ByRef Buffer As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    If IsMissing(Buffer) Then
        Read = ReadOneChar
    Else
        Read = ReadToBuffer(Buffer, Index, Count)
    End If
End Function

''
' Returns the next character to be decoded from the byte stream. The position
' in the stream is not advanced after the read.
'
' @return The next character to be Read, or -1 if no more characters are in the stream.
' @remarks This function requires the stream to support Seeking. If it does not,
' then -1 will always be returned.
'
Public Function PeekChar() As Long
    VerifyIsOpen
    
    If mStream.CanSeek Then
        Dim OldPosition As Long
        OldPosition = mStream.Position
        PeekChar = Read
        mStream.Position = OldPosition
    Else
        PeekChar = NoChars
    End If
End Function

''
' Reads the next byte in the stream, and advances the position one byte.
'
' @return The next byte in the stream.
Public Function ReadByte() As Byte
    FillBuffer corSizeOfByte
    ReadByte = mBuffer(0)
End Function

''
' Reads the next decode character in the stream, and advances the position
' the number of bytes requirece to assemble a single character.
'
' @return The next character in the stream.
' @remarks If no characters are left in the stream, and EndOfStreamException is thrown.
'
Public Function ReadChar() As Long
    ReadChar = Read
    If ReadChar = NoChars Then _
        Throw New EndOfStreamException
End Function

''
' Reads the specified number of bytes from the stream and returns them in an array.
'
' @param Count The number of bytes to read from the stream.
' @return A byte array containing the bytes read from the stream.
' @remarks If not enough bytes are in the stream, then the remaining
' bytes are returned. If there are no bytes in the stream, then an
' empty array is returned, not a null array.
'
Public Function ReadBytes(ByVal Count As Long) As Byte()
    VerifyIsOpen
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    
    Dim Bytes() As Byte
    If Count > 0 Then
        ReDim Bytes(0 To Count - 1)
        
        Dim NumberOfBytesRead As Long
        NumberOfBytesRead = mStream.ReadBlock(Bytes, 0, Count)
        
        If NumberOfBytesRead = 0 Then Exit Function
        
        If NumberOfBytesRead < Count Then
            ReDim Preserve Bytes(0 To NumberOfBytesRead - 1)
        End If
    Else
        Bytes = Cor.NewBytes()
    End If
    
    ReadBytes = Bytes
End Function

''
' Returns a specified number of characters decoded from the stream.
'
' @param Count The number of characters to decode from the stream.
' @return The array of characters decoded from the stream.
' @remarks If not enough characters could be decoded from the stream,
' then only the remaining characters are returned. If no characters
' could be decoded, then an empty array is returned, not a null array.
'
Public Function ReadChars(ByVal Count As Long) As Integer()
    VerifyIsOpen
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    
    Dim Chars() As Integer
    If Count > 0 Then
        ReDim Chars(0 To Count - 1)
        
        Dim NumberOfCharsRead As Long
        NumberOfCharsRead = ReadCharBytes(Chars, 0, Count)
        If NumberOfCharsRead = 0 Then _
            Throw New EndOfStreamException
        
        If NumberOfCharsRead < Count Then
            ReDim Preserve Chars(0 To NumberOfCharsRead - 1)
        End If
    Else
        Chars = Cor.NewIntegers()
    End If
    
    ReadChars = Chars
End Function

''
' Returns a boolean from the stream.
'
' @return A boolean value.
' @remarks A single byte is read from the stream. If the byte
' is non-zero, then True is returned, otherwise False.
'
Public Function ReadBoolean() As Boolean
    ReadBoolean = (ReadByte <> 0)
End Function

''
' Reads 4 bytes from the stream and returns them as a Long datatype.
'
' @return A Long datatype read from the stream.
'
Public Function ReadInt32() As Long
    FillBuffer corSizeOfLong
    ReadInt32 = AsLong(mBuffer(0))
End Function

''
' Reads 2 bytes from the stream and returns them as an Integer datatype.
'
' @return An Integer datatype read from the stream.
'
Public Function ReadInt16() As Integer
    FillBuffer corSizeOfInteger
    ReadInt16 = AsWord(mBuffer(0))
End Function

''
' Reads 8 bytes from the stream and returns them as a Double datatype.
'
' @return A Double datatype read from the stream.
'
Public Function ReadDouble() As Double
    FillBuffer corSizeOfDouble
    ReadDouble = AsDouble(mBuffer(0))
End Function

''
' Reads 8 bytes from the stream and returns them as a Date datatype.
'
' @return A Date datatype read from the stream.
' @remarks It is possible to read a combination of 8 bytes that
' convert into an invalid Date. This will cause and Overflow error.
'
Public Function ReadDate() As Date
    Dim PotentialDate As Double
    PotentialDate = ReadDouble
    
    If PotentialDate < MinDateAsDouble Or PotentialDate > MaxDateAsDouble Then _
        Throw New OverflowException
        
    ReadDate = PotentialDate
End Function

''
' Reads 8 bytes from the stream and returns them as a Currency datatype.
'
' @return A Currency datatype.
'
Public Function ReadCurrency() As Currency
    FillBuffer corSizeOfCurrency
    ReadCurrency = AsCurr(mBuffer(0))
End Function

''
' Reads 4 bytes from the stream and returns them as a Single datatype.
'
' @return A Single datatype.
'
Public Function ReadSingle() As Single
    FillBuffer corSizeOfSingle
    ReadSingle = AsSingle(mBuffer(0))
End Function

''
' Reads 16 bytes from the stream and returns them as a Variant Decimal datatype.
'
' @return A Variant Decimal datatype.
' @remarks <p>The Decimal datatype is created from 16 bytes with the same
' layout as used in .NET. The layout is different than that of VB.<br><br>
' VB Decimal layout<br>
' bytes 0-1: Variant datatype information<br>
' bytes 2: precision<br>
' bytes 3: sign (&h80 is negative)<br>
' bytes 4-7: the 32 highest bits in the 96bit value<br>
' bytes 8-11: the 32 lowest bits in the 96bit value<br>
' bytes 12-15: the 32 middle bits in the 96bit value<br>
' <br>
' .NET Decimal layout<br>
' bytes 0-3: the 32 lowest bits in the 96bit value<br>
' bytes 4-7: the 32 middle bits in the 96bit value<br>
' bytes 8-11: the 32 highest bits in the 96bit value<br>
' bytes 12-13: unused (zero)<br>
' bytes 14: precision<br>
' bytes 15: sign (&h80 is negative)<br><br>
' If the sign byte is non-zero, then the Decimal will be set as negative.<br>
'
Public Function ReadDecimal() As Variant
    FillBuffer corSizeOfDecimal
    ReadDecimal = BitConverter.ToDecimal(mBuffer, 0)
End Function

''
' Reads a String from the stream.
'
' @return The string datatype.
' @remarks The string is stored in the stream with the number of encoded bytes preceding
' the actual string data. The bytes are not the typical 4 bytes as is used by VB when
' representing the length of the string. The bytes represent a value encoded as
' 7bits per byte. While the next read byte has the high bit set (&h80), then the
' following byte is also part of the length value. For each byte read, first 7 bits of
' that byte is pushed out 7 bits multiplied by the current byte number - 1 in the sequence.
' This is not the normal shifting of the current sum of the values. Each read byte
' must be shifted left individually, as each byte represents a higher set of bits
' in the resulting number.<br><br>
' We AND the byte with &h3f because we only want the lower 7 bits.<br>
' byte 1: (byte and &h3f) << shifted 0 bits added to the sum<br>
' byte 2: (byte and &h3f) << shifted 7 bits added to the sum<br>
' byte 3: (byte and &H3f) << shifted 14 bits added to the sum<br><br>
' .. this continues until a byte is less than 128. At which point, it is shifted and summed like the
' rest, but no more bytes are to be read in. The sum now contains the number of bytes to
' be read in and decoded into the string. The same Encoding type must be used to retrieve
' the correct string value. Using a different Encoding method will create the wrong string.
'
' <p>Once the length is determined, then enough bytes are read in that can be decoded
' to the string using the current Encoding system.</p>
'
Public Function ReadString() As String
    VerifyIsOpen
    
    Dim EncodedStringLength As Long
    EncodedStringLength = ReadStringLength
    
    If EncodedStringLength > 0 Then
        ReadString = ReadEncodedString(EncodedStringLength)
    End If
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
' @IObject
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
' @IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
' @IObject
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    ReDim mBuffer(0 To DefaultCapacity - 1)
End Sub

Friend Sub Init(ByVal Stream As Stream, ByVal Encoding As Encoding)
    If Stream Is Nothing Then _
        Error.ArgumentNull "Stream", ArgumentNull_Stream
    If Not Stream.CanRead Then _
        Error.Argument Argument_ReadableStreamRequired
    
    Set mStream = Stream
    If Encoding Is Nothing Then
        Set mEncoding = New UTF8Encoding
    Else
        Set mEncoding = Encoding
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillBuffer(ByVal Count As Long)
    VerifyIsOpen
    
    If Count = 1 Then
        Dim SingleByte As Long
        SingleByte = mStream.ReadByte
        If SingleByte < 0 Then _
            Throw New EndOfStreamException
        
        mBuffer(0) = SingleByte
    Else
        If mStream.ReadBlock(mBuffer, 0, Count) < Count Then
            Throw New EndOfStreamException
        End If
    End If
End Sub

Private Sub VerifyIsOpen()
    If mStream Is Nothing Then
        Error.ObjectDisposed ObjectDisposed_FileNotOpen
    End If
End Sub

Private Function ReadOneChar() As Long
    VerifyIsOpen
    If ReadCharBytes(mSingleChar, 0, 1) = 0 Then
        ReadOneChar = NoChars
    Else
        AsWord(ReadOneChar) = mSingleChar(0)
    End If
End Function

Private Function ReadToBuffer(ByRef Buffer As Variant, ByRef Index As Variant, ByRef Count As Variant) As Long
    Dim BufferPtr As Long
    BufferPtr = CorArray.ArrayPointer(Buffer)
    Validation.CheckValidSingleDimArrayPtr BufferPtr, Parameter_Buffer
    
    Dim Range As ListRange
    Range = OptionalRange(Index, Count, LBound(Buffer), CorArray.LengthFirstDim(Buffer))
    
    If Range.Index < LBound(Buffer) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count - 1 > UBound(Buffer) Then _
        Error.Argument Argument_InvalidOffLen
    
    Select Case VarType(Buffer)
        Case corByteArray
            ReadToBuffer = ReadToByteBuffer(BufferPtr, Range)
        Case corIntegerArray
            ReadToBuffer = ReadToCharBuffer(BufferPtr, Range)
        Case Else
            Error.Argument Argument_UnsupportedArray, "Buffer"
    End Select
End Function

Private Function ReadToByteBuffer(ByVal BufferPtr As Long, ByRef BufferRange As ListRange) As Long
    Dim Buffer() As Byte
    
    On Error GoTo Catch
    SAPtr(Buffer) = BufferPtr
    ReadToByteBuffer = mStream.ReadBlock(Buffer, BufferRange.Index, BufferRange.Count)
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw Err
Finally:
    SAPtr(Buffer) = vbNullPtr
    Return
End Function

Private Function ReadToCharBuffer(ByVal BufferPtr As Long, ByRef BufferRange As ListRange) As Long
    Dim Buffer() As Integer
    
    On Error GoTo Catch
    SAPtr(Buffer) = BufferPtr
    ReadToCharBuffer = ReadCharBytes(Buffer, BufferRange.Index, BufferRange.Count)
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw Err
Finally:
    SAPtr(Buffer) = vbNullPtr
    Return
End Function

Private Function ReadCharBytes(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    Dim SingleByte          As Long
    Dim BufferPos           As Long
    Dim TotalCharsRead      As Long
    Dim NumberOfCharsRead   As Long
    
    Do While TotalCharsRead < Count
        SingleByte = mStream.ReadByte
        
        If SingleByte < 0 Then
            Exit Do
        End If
        
        mBuffer(BufferPos) = SingleByte
        BufferPos = BufferPos + 1
        
        NumberOfCharsRead = mEncoding.GetCharsEx(mBuffer, 0, BufferPos, Chars, Index)
        
        If NumberOfCharsRead > 0 Then
            TotalCharsRead = TotalCharsRead + NumberOfCharsRead
            Index = Index + NumberOfCharsRead
            BufferPos = 0
            NumberOfCharsRead = 0
        End If
    Loop
    
    ReadCharBytes = TotalCharsRead
End Function

Private Function ReadStringLength() As Long
    Const BitsShiftedAtATime    As Long = 7
    Const BitsToShiftMask       As Long = &H7F
    
    Dim StringLength    As Long
    Dim SingleByte      As Long
    Dim ShiftCount      As Long
    
    SingleByte = mStream.ReadByte
    
    Do While SingleByte >= &H80
        StringLength = StringLength + Helper.ShiftLeft(SingleByte And BitsToShiftMask, ShiftCount)
        SingleByte = mStream.ReadByte
        ShiftCount = ShiftCount + BitsShiftedAtATime
    Loop
    
    If SingleByte > 0 Then
        StringLength = StringLength + Helper.ShiftLeft(SingleByte, ShiftCount)
    End If
    
    ReadStringLength = StringLength
End Function

Private Function ReadEncodedString(ByVal StringLength As Long) As String
    Dim EncodedBytes() As Byte
    ReDim EncodedBytes(0 To StringLength - 1)
    
    Dim NumberOfBytesRead As Long
    NumberOfBytesRead = mStream.ReadBlock(EncodedBytes, 0, StringLength)
    
    If NumberOfBytesRead < StringLength Then _
        Throw New EndOfStreamException
        
    ReadEncodedString = mEncoding.GetString(EncodedBytes)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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
