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

Private Const DEF_CAPACITY  As Long = 16
Private Const MIN_DATE      As Double = -657434#
Private Const MAX_DATE      As Double = 2958465.99998843


Private mStream         As Stream
Private mEncoding       As Encoding
Private mIsOpen         As Boolean
Private mBuffer()       As Byte
Private mCapacity       As Long
Private mPtrBuffer      As Long
Private mOneChar(0)     As Integer


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
    mStream.CloseStream
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
    Select Case CLng(IsMissing(Buffer)) + CLng(IsMissing(Index)) + CLng(IsMissing(Count))
        Case -3:    Read = ReadOneChar
        Case 0:     Read = ReadBuffer(Buffer, Index, Count)
        Case Else:  Throw Cor.NewInvalidOperationException("All parameters must be supplied or missing.")
    End Select
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
        PeekChar = -1
    End If
End Function

''
' Reads the next byte in the stream, and advances the position one byte.
'
' @return The next byte in the stream.
Public Function ReadByte() As Byte
    FillBuffer 1
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
    If ReadChar = -1 Then _
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
        Throw Cor.NewArgumentOutOfRangeException("Count", Message:=Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    
    Dim Ret() As Byte
    If Count > 0 Then
        ReDim Ret(0 To Count - 1)
        
        Dim BytesRead As Long
        BytesRead = mStream.ReadBlock(Ret, 0, Count)
        
        If BytesRead = 0 Then Exit Function
        
        If BytesRead < Count Then
            ReDim Preserve Ret(0 To BytesRead - 1)
        End If
    Else
        Ret = Cor.NewBytes()
    End If
    
    ReadBytes = Ret
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
        Throw Cor.NewArgumentOutOfRangeException("Count", Message:=Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    
    Dim Ret() As Integer
    If Count > 0 Then
        ReDim Ret(0 To Count - 1)
        
        Dim CharsRead As Long
        CharsRead = ReadCharBytes(Ret, 0, Count)
        If CharsRead = 0 Then _
            Throw New EndOfStreamException
        
        If Count < CharsRead Then
            ReDim Preserve Ret(0 To CharsRead - 1)
        End If
    Else
        Ret = Cor.NewIntegers()
    End If
    
    ReadChars = Ret
End Function

''
' Returns a boolean from the stream.
'
' @return A boolean value.
' @remarks A single byte is read from the stream. If the byte
' is non-zero, then True is returned, otherwise False.
'
Public Function ReadBoolean() As Boolean
    VerifyIsOpen
    ReadBoolean = (ReadByte <> 0)
End Function

''
' Reads 4 bytes from the stream and returns them as a Long datatype.
'
' @return A Long datatype read from the stream.
'
Public Function ReadLong() As Long
    FillBuffer 4
    ReadLong = MemLong(mPtrBuffer)
End Function

''
' Reads 2 bytes from the stream and returns them as an Integer datatype.
'
' @return An Integer datatype read from the stream.
'
Public Function ReadInteger() As Integer
    FillBuffer 2
    ReadInteger = MemWord(mPtrBuffer)
End Function

''
' Reads 8 bytes from the stream and returns them as a Double datatype.
'
' @return A Double datatype read from the stream.
'
Public Function ReadDouble() As Double
    FillBuffer 8
    ReadDouble = AsDouble(ByVal mPtrBuffer)
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
    
    If PotentialDate < MIN_DATE Or PotentialDate > MAX_DATE Then _
        Throw New OverflowException
        
    ReadDate = PotentialDate
End Function

''
' Reads 8 bytes from the stream and returns them as a Currency datatype.
'
' @return A Currency datatype.
'
Public Function ReadCurrency() As Currency
    FillBuffer 8
    ReadCurrency = AsCurr(ByVal mPtrBuffer)
End Function

''
' Reads 4 bytes from the stream and returns them as a Single datatype.
'
' @return A Single datatype.
'
Public Function ReadSingle() As Single
    FillBuffer 4
    ReadSingle = AsSingle(ByVal mPtrBuffer)
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
    FillBuffer 16
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
        Dim EncodedBytes() As Byte
        ReDim EncodedBytes(0 To EncodedStringLength - 1)
        
        Dim BytesRead As Long
        BytesRead = mStream.ReadBlock(EncodedBytes, 0, EncodedStringLength)
        
        If BytesRead < EncodedStringLength Then _
            Throw New EndOfStreamException
            
        ReadString = mEncoding.GetString(EncodedBytes)
    End If
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
' @IObject
'
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
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
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
' @IObject
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Stream As Stream, ByVal Encoding As Encoding)
    If Stream Is Nothing Then _
        Throw Cor.NewArgumentNullException("Stream", Resources.GetString(ArgumentNull_Stream))
    If Not Stream.CanRead Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_ReadableStreamRequired), "Stream")
    
    Set mStream = Stream
    If Encoding Is Nothing Then
        Set mEncoding = New UTF8Encoding
    Else
        Set mEncoding = Encoding
    End If
    mIsOpen = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyIsOpen()
    If Not mIsOpen Then Throw Cor.NewObjectDisposedException("BinaryReader", "Cannot read from a closed BinaryReader.")
End Sub

Private Sub FillBuffer(ByVal Count As Long)
    VerifyIsOpen
    
    ' optimize for reading a single byte
    Dim num As Long
    If Count = 1 Then
        num = mStream.ReadByte
        If num >= 0 Then
            mBuffer(0) = num
            Exit Sub
        End If
    Else
        num = mStream.ReadBlock(mBuffer, 0, Count)
    End If
    
    If num < Count Then _
        Throw New EndOfStreamException
End Sub

Private Function ReadStringLength() As Long
    Dim Ret         As Long
    Dim num         As Long
    Dim shiftCount  As Long
    
    num = mStream.ReadByte
    Do While num >= &H80
        Ret = Ret + Helper.ShiftLeft(num And &H7F, shiftCount)
        num = mStream.ReadByte
        shiftCount = shiftCount + 7
    Loop
    If num > 0 Then
        ReadStringLength = Ret + Helper.ShiftLeft(num, shiftCount)
    Else
        ReadStringLength = Ret
    End If
End Function

Private Function ReadCharBytes(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    Dim Result As Long
    
    Result = VerifyArrayRange(SAPtr(Chars), Index, Count)
    If Result <> NO_ERROR Then _
        ThrowArrayRangeException Result, "Chars", Index, "Index", Count, "Count"
    
    EnsureCapacity mEncoding.GetMaxByteCount(Count)
    
    Dim ByteRead    As Long
    Dim Pos         As Long
    Dim NumRead     As Long
    Do While NumRead < Count
        ByteRead = mStream.ReadByte
        If ByteRead < 0 Then Exit Do
        mBuffer(Pos) = ByteRead
        Pos = Pos + 1
        NumRead = NumRead + mEncoding.GetCharsEx(mBuffer, 0, Pos, Chars, Index)
    Loop
    ReadCharBytes = Pos
End Function

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    If RequiredCapacity <= mCapacity Then _
        Exit Sub
    
    Dim NewCapacity As Long
    NewCapacity = mCapacity * 2
    If NewCapacity < RequiredCapacity Then
        NewCapacity = RequiredCapacity
    End If
    
    ReDim Preserve mBuffer(0 To NewCapacity - 1)
    mCapacity = NewCapacity
    mPtrBuffer = VarPtr(mBuffer(0))
End Sub

Private Function ReadBuffer(ByRef Buffer As Variant, ByRef Index As Variant, ByRef Count As Variant) As Long
    Dim ElemIndex As Long
    Dim ElemCount As Long
    
    ElemIndex = GetOptionalLong(Index, 0)
    ElemCount = GetOptionalLong(Count, 0)
    
    Dim Bytes() As Byte
    Dim Chars() As Integer

    On Error GoTo errTrap
    Select Case VarType(Buffer)
        Case vbByteArray
            SAPtr(Bytes) = GetArrayPointer(Buffer, True)
            ReadBuffer = mStream.ReadBlock(Bytes, ElemIndex, ElemCount)
        Case vbIntegerArray
            SAPtr(Chars) = GetArrayPointer(Buffer, True)
            ReadBuffer = ReadCharBytes(Chars, ElemIndex, ElemCount)
        Case Else
            On Error GoTo 0
            Throw Cor.NewArgumentException("A Byte or Integer array is required.", "Buffer")
    End Select

errTrap:
    SAPtr(Bytes) = 0
    SAPtr(Chars) = 0
    Throw
End Function

Private Function ReadOneChar() As Long
    VerifyIsOpen
    If ReadCharBytes(mOneChar, 0, 1) = 0 Then
        ReadOneChar = -1
    Else
        AsWord(ReadOneChar) = mOneChar(0)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    EnsureCapacity DEF_CAPACITY
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
