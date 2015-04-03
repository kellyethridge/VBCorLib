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

Private Const DefaultCapacity  As Long = 16
Private Const MinDateAsDouble  As Double = -657434#
Private Const MaxDateAsDouble  As Double = 2958465.99998843


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
    If IsMissing(Buffer) Then
        Read = ReadOneChar
    Else
        Dim BufferPtr As Long: BufferPtr = Require.ValidOneDimensionArray(Buffer, Param.BufferParam, ArgumentNull.NullBuffer)
        Dim BufferIndex As Long: BufferIndex = GetOptionalLong(Index, LBound(Buffer))
        Dim BufferCount As Long: BufferCount = GetOptionalLong(Count, ArrayLength(Buffer))
        
        Require.ThatArgument BufferIndex >= LBound(Buffer), Param.Index, ArgumentOutOfRange.ArrayLB
        Require.ThatArgument BufferCount >= 0, Param.Count, ArgumentOutOfRange.NeedNonNegNum
        Require.That BufferIndex + BufferCount - 1 <= UBound(Buffer), Argument.InvalidOffLen
        
        Select Case VarType(Buffer)
            Case vbByteArray
                Read = ReadToByteBuffer(BufferPtr, BufferIndex, BufferCount)
            Case vbIntegerArray
                Read = ReadToCharBuffer(BufferPtr, BufferIndex, BufferCount)
            Case Else
                Throw Cor.NewArgumentException("A Byte or Integer array is required.", GetParameter(Param.BufferParam))
        End Select
    End If
End Function

Private Function ReadToByteBuffer(ByVal BufferPtr As Long, ByVal Index As Long, ByVal Count As Long) As Long
    Dim Buffer() As Byte
    
    On Error GoTo errTrap
    SAPtr(Buffer) = BufferPtr
    ReadToByteBuffer = mStream.ReadBlock(Buffer, Index, Count)
    
errTrap:
    SAPtr(Buffer) = vbNullPtr
End Function

Private Function ReadToCharBuffer(ByVal BufferPtr As Long, ByVal Index As Long, ByVal Count As Long) As Long
    Dim Buffer() As Integer
    
    On Error GoTo errTrap
    SAPtr(Buffer) = BufferPtr
    ReadToCharBuffer = ReadCharBytes(Buffer, Index, Count)
    
errTrap:
    SAPtr(Buffer) = vbNullPtr
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
    Require.ThatArgument Count >= 0, Param.Count, ArgumentOutOfRange.NeedNonNegNum
    VerifyIsOpen
    
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
    Require.ThatArgument Count >= 0, Param.Count, ArgumentOutOfRange.NeedNonNegNum
    VerifyIsOpen
    
    Dim Result() As Integer
    If Count > 0 Then
        ReDim Result(0 To Count - 1)
        
        Dim CharsRead As Long
        CharsRead = ReadCharBytes(Result, 0, Count)
        If CharsRead = 0 Then _
            Throw New EndOfStreamException
        
        If CharsRead < Count Then
            ReDim Preserve Result(0 To CharsRead - 1)
        End If
    Else
        Result = Cor.NewIntegers()
    End If
    
    ReadChars = Result
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
Public Function ReadInt32() As Long
    FillBuffer 4
    ReadInt32 = AsLong(mBuffer(0))
End Function

''
' Reads 2 bytes from the stream and returns them as an Integer datatype.
'
' @return An Integer datatype read from the stream.
'
Public Function ReadInt16() As Integer
    FillBuffer 2
    ReadInt16 = AsWord(mBuffer(0))
End Function

''
' Reads 8 bytes from the stream and returns them as a Double datatype.
'
' @return A Double datatype read from the stream.
'
Public Function ReadDouble() As Double
    FillBuffer 8
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
    FillBuffer 8
    ReadCurrency = AsCurr(mBuffer(0))
End Function

''
' Reads 4 bytes from the stream and returns them as a Single datatype.
'
' @return A Single datatype.
'
Public Function ReadSingle() As Single
    FillBuffer 4
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
    Require.ObjectNotNothing Stream, Param.StreamParam, ArgumentNull.NullStream
    Require.That Stream.CanRead, Argument.ReadableStreamRequired, Param.StreamParam
    
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
    If Not mIsOpen Then
        Throw Cor.NewObjectDisposedException("BinaryReader", "Cannot read from a closed BinaryReader.")
    End If
End Sub

Private Sub FillBuffer(ByVal Count As Long)
    VerifyIsOpen
    
    ' optimize for reading a single byte
    Dim ByteRead As Long
    If Count = 1 Then
        ByteRead = mStream.ReadByte
        If ByteRead < 0 Then _
            Throw New EndOfStreamException
        
        mBuffer(0) = ByteRead
        Exit Sub
    End If
    
    If mStream.ReadBlock(mBuffer, 0, Count) < Count Then _
        Throw New EndOfStreamException
End Sub

Private Function ReadStringLength() As Long
    Dim Result      As Long
    Dim ByteRead    As Long
    Dim ShiftCount  As Long
    
    ByteRead = mStream.ReadByte
    Do While ByteRead >= &H80
        Result = Result + Helper.ShiftLeft(ByteRead And &H7F, ShiftCount)
        ByteRead = mStream.ReadByte
        ShiftCount = ShiftCount + 7
    Loop
    If ByteRead > 0 Then
        ReadStringLength = Result + Helper.ShiftLeft(ByteRead, ShiftCount)
    Else
        ReadStringLength = Result
    End If
End Function

Private Function ReadCharBytes(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    Dim ByteRead    As Long
    Dim BufferPos   As Long
    Dim NumRead     As Long
    Dim CharsRead   As Long
    
    Do While NumRead < Count
        ByteRead = mStream.ReadByte
        If ByteRead < 0 Then Exit Do
        mBuffer(BufferPos) = ByteRead
        BufferPos = BufferPos + 1
        
        CharsRead = mEncoding.GetCharsEx(mBuffer, 0, BufferPos, Chars, Index)
        
        If CharsRead > 0 Then
            NumRead = NumRead + CharsRead
            Index = Index + CharsRead
            BufferPos = 0
            CharsRead = 0
        End If
    Loop
    
    ReadCharBytes = NumRead
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
    ReDim mBuffer(0 To DefaultCapacity - 1)
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
