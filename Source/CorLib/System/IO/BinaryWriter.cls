VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryWriter"
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
' Module: BinaryWriter
'

''
' Provides methods to write typed information to a stream.
'
' @remarks
' <p>The datatypes that are written to the stream are written in their binary
' form. The value are written as they are represented in memory. For example, as datatype
' of vbLong is written as 4 bytes, regardless of the actual value of the vbLong
' variable.</p>
'
' @see Constructors
' @see BinaryReader
'
Option Explicit
Implements IObject

Private mOutStream  As Stream
Private mEncoding   As Encoding
Private mIsOpen     As Boolean
Private mBuffer()   As Byte
Private mBufferSA   As SafeArray1d
Private mBytes()    As Byte
Private mOneChar(0) As Integer


''
' Returns the underlying stream this writer is writing to.
'
' @return The underlying stream.
' @remarks The stream is flushed before it is returned.
'
Public Property Get BaseStream() As Stream
    Flush
    Set BaseStream = mOutStream
End Property

''
' Closes the writer, flushing any data.
'
Public Sub CloseWriter()
    If mIsOpen Then
        Flush
        mOutStream.CloseStream
        mIsOpen = False
    End If
End Sub

''
' Tells the underlying stream to store its data to its repository.
'
Public Sub Flush()
    mOutStream.Flush
End Sub

''
' Sets the underlying stream's current position.
'
' @param Offset The number of bytes to move the position.
' @param Origin The starting point to move from in the stream.
' @return The final position in the stream.
'
Public Function SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency
    mOutStream.SeekPosition Offset, Origin
End Function

''
' Writes the binary representation of a datatype to the underlying stream.
'
' @param value The datatype to be written.
' @param Index The starting index of arrays or string characters to being writing from.
' @param Count The number of elements or characters to be written to the stream.
' @remarks <i>Index</i> and <i>vCount</i> are only valid when writing Byte arrays,
' Integer array, and Strings. For all other types, these parameters are ignored.
'<p>Simple datatypes such as vbLong and vbDouble are written out exactly as they
' appear in memory. Each individual byte of that datatype is written. For example, a vbLong
' is 4 bytes in memory, even if its value is 0. The 4 bytes are written to the stream, not
' one byte with a value of 0. A vbDouble is 8 bytes, so all 8 bytes are written as is to
' the stream.</p>
' <p>The one exception is when writing a String value. Both the number of bytes from an encoded
' string and the encoded bytes are written, however, the length is not written as the individual bytes, like the
' other datatypes. The length of the string is written as a series of 7bit values. Each byte
' for the length is calculated by starting out with the actual value, then ANDing it with &h7F
' to take the first 7 bits. The value is ORed with &h80 to signify there are more bytes to
' follow for the length. Once that byte is written, the original value is shifted right
' 7 bits and the process starts over until the value reaches less than &h80. The final byte
' value is not ORed with &h80, it is written as is without the high bit set. When reading
' the bytes back, the byte without the high bit set will signal as the last byte in the value.</p><br>
' <pre>
'    Do While value > &H80
'        WriteValue CByte((value And &H7F) Or &H80)
'        value = (value And &HFFFFFF80) \ &H80
'    Loop
'    WriteValue CByte(value)
' </pre><br>
' The characters of the string are written as the encoded bytes generated by the current Encoding.<br>
' <p>Byte arrays are written as is. There is no description written along with the byte values.
' The number of bytes will have to be known in advance or manually stored in the stream.</p>
' <p>Integer arrays are treated as characters and are converted to the encoded bytes using the
' current encoding. Like the byte array, the number of bytes is not stored with the encoded
' characters and will have to be manually stored.</p>
' <br>
' <p>The Decimal datatype is converted to 16 bytes with the same
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
' bytes 15: sign (&h80 is negative)<br>
'
Public Sub WriteValue(ByRef Value As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    VerifyIsOpen

    If IsArray(Value) Then
        WriteArray Value, Index, Count
    Else
        WriteIntrinsic Value, Index, Count
    End If
End Sub

''
' Writes a character to the underlying stream.
'
' @param Value The character to be written. This can be a string or numeric value.
' @remarks The character is encoded and the byte array is written to the Stream.
'
Public Sub WriteChar(ByRef Value As Variant)
    Select Case VarType(Value)
        Case vbString
            mOneChar(0) = AscW(CStr(Value))
        Case vbLong, vbInteger, vbByte
            mOneChar(0) = CInt(Value)
        Case Else
            Throw Cor.NewArgumentException("Value must be a String or Integer value.", "Value")
    End Select
    
    WriteChars mOneChar
End Sub

''
' Writes a boolean to the stream as a one-byte value and advances the stream position one byte.
'
' @param Value The boolean value to be written to the stream.
' @remarks When writting boolean values to the stream, a 1 is written for True and 0 for false.
'
Public Sub WriteBoolean(ByVal Value As Boolean)
    VerifyIsOpen
    
    If Value Then
        mOutStream.WriteByte 1
    Else
        mOutStream.WriteByte 0
    End If
End Sub

''
' Writes a one-byte unsigned integer value to the stream and advances the stream position one byte.
'
' @param Value The one-byte unsigned integer to write.
'
Public Sub WriteByte(ByVal Value As Byte)
    VerifyIsOpen
    mOutStream.WriteByte Value
End Sub

''
' Writes a four-byte integer to the stream and advances the stream position four bytes.
'
' @param Value The four-byte integer to write.
' @remarks The four byte value is written in Little Endian format.
'
Public Sub WriteInt32(ByVal Value As Long)
    VerifyIsOpen
    mOutStream.WriteBlock AsQByteArr(Value).Bytes, 0, 4
End Sub

''
' Writes a two-byte integer value to the stream and advances the stream position two bytes.
'
' @param Value The two-byte integer to write.
' @remarks The two-byte value is written in Little Endian format.
'
Public Sub WriteInt16(ByVal Value As Integer)
    VerifyIsOpen
    mOutStream.WriteBlock AsDByteArr(Value).Bytes, 0, 2
End Sub

''
' Writes a four-byte floating point value to the stream and advances the stream position by four bytes.
'
' @param Value The four-byte floating point value to write.
' @remarks The four-byte value is written in Little Endian format.
'
Public Sub WriteSingle(ByVal Value As Single)
    VerifyIsOpen
    mOutStream.WriteBlock AsQByteArr(Value).Bytes, 0, 4
End Sub

''
' Writes an eight-byte floating point value to the stream and advances the stream position by eight bytes.
'
' @param Value The eight-byte floating point value to write.
' @remarks The eight-byte value is written in Little Endian format.
'
Public Sub WriteDouble(ByVal Value As Double)
    VerifyIsOpen
    mOutStream.WriteBlock AsOByteArr(Value).Bytes, 0, 8
End Sub

''
' Writes a Date value to the stream and advances the stream position by eight bytes.
'
' @param Value The eight-byte VBA Date value to write.
' @remarks The eight-byte value is written in the same manner as a Double.
'
Public Sub WriteDate(ByVal Value As Date)
    VerifyIsOpen
    mOutStream.WriteBlock AsOByteArr(Value).Bytes, 0, 8
End Sub

''
' Writes a Currency value to the stream and advances the stream position by sixteen bytes.
'
' @param Value The currency value to write.
' @remarks The binary representation of the value is written in Little Endian.
'
Public Sub WriteCurrency(ByVal Value As Currency)
    VerifyIsOpen
    mOutStream.WriteBlock AsOByteArr(Value).Bytes, 0, 8
End Sub

''
' Writes a decimal value to the stream and andvances the stream position by sixteen bytes.
'
' @param Value The decimal value to write.
' @remarks <p>The decimal datatype is converted to sixteen bytes with the same
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
' bytes 15: sign (&h80 is negative)<br>
'
Public Sub WriteDecimal(ByRef Value As Variant)
    VerifyIsOpen
    mOutStream.WriteBlock BitConverter.GetBytes(CDec(Value)), 0, 16
End Sub

''
' Writes a length-prefixed string to the stream using the encoding in the BinaryWriter and advances the stream position
' in accordance to the number of characters the encoding produces.
'
' @param Value The string to encode and write.
' @remarks The number of bytes from an encoded string and the encoded bytes are written,
' however, the length is not written as the individual bytes, like the other datatypes.
' The length of the string is written as a series of 7bit values. Each byte
' for the length is calculated by starting out with the actual value, then ANDing it with &h7F
' to take the first 7 bits. The value is ORed with &h80 to signify there are more bytes to
' follow for the length. Once that byte is written, the original value is shifted right
' 7 bits and the process starts over until the value reaches less than &h80. The final byte
' value is not ORed with &h80, it is written as is without the high bit set. When reading
' the bytes back, the byte without the high bit set will signal the last byte in the value.</p><br>
' <pre>
'    Do While LengthRemaining > &H80
'        WriteByte (LengthRemaining And &H7F) Or &H80
'        LengthRemaining = (LengthRemaining And &HFFFFFF80) \ &H80
'    Loop
'    WriteByte LengthRemaining
' </pre><br>
' The characters of the string are written as the encoded bytes generated by the current Encoding.
'
Private Sub WriteString(ByRef Value As String)
    VerifyIsOpen
    WriteStringLength mEncoding.GetByteCount(Value)
    WriteBytes mEncoding.GetBytes(Value)
End Sub

''
' Writes a section of a byte array to the stream and advances the stream position by the number of bytes written.
'
' @param Buffer The array of bytes to write.
' @param Index The starting index within the buffer to begin writing.
' @param Count The number of bytes to write.
' @remarks If the index and count are not specified, the lowerbound and length of the array will be used.
'
Public Sub WriteBytes(ByRef Buffer() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    VerifyIsOpen
    Require.NotNullSingleDimensionArray Buffer, Parameter_Buffer
    
    Dim Range As ListRange
    Range = GetOptionalRange(Index, Count, LBound(Buffer), ArrayLength(Buffer))
    Validate.ArrayRange Buffer, Range
    mOutStream.WriteBlock Buffer, Range.Index, Range.Count
End Sub

''
' Writes a section of a character array using the encoding in the BinaryWriter and advances the stream position
' in accordance to the number of characters the encoding produces.
'
' @param Chars The characters to encode and write.
' @param Index The starting index within the characters to begin writing.
' @param Count The number of characters to write.
' @remarks If the index and count are not specified, the lowerbound and length of the array will be used.
'
Public Sub WriteChars(ByRef Chars() As Integer, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    VerifyIsOpen
    Require.NotNullSingleDimensionArray Chars, Parameter_Chars
    
    Dim Range As ListRange
    Range = GetOptionalRange(Index, Count, LBound(Chars), ArrayLength(Chars))
    Validate.ArrayRange Chars, Range
    
    Dim Bytes() As Byte
    Bytes = mEncoding.GetBytes(Chars, Range.Index, Range.Count)
    mOutStream.WriteBlock Bytes, 0, ArrayLength(Bytes)
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
' @see IObject
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
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
' @see IObject
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Output As Stream, ByVal Encoding As Encoding)
    Require.NotNothing Output, Parameter_Output
    Require.That Output.CanWrite, Argument_StreamNotWritable
    
    Set mOutStream = Output
    If Encoding Is Nothing Then
        Set mEncoding = Cor.NewUTF8Encoding(False, True)
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
        Throw Cor.NewObjectDisposedException("BinaryWriter", "Cannot write to a closed BinaryWriter.")
    End If
End Sub

Private Sub WriteStringLength(ByVal Length As Long)
    Dim LengthRemaining As Long
    LengthRemaining = Length
    
    Do While LengthRemaining > &H80
        WriteByte (LengthRemaining And &H7F) Or &H80
        LengthRemaining = (LengthRemaining And &HFFFFFF80) \ &H80
    Loop
    WriteByte LengthRemaining
End Sub

Private Sub WriteIntrinsic(ByRef Value As Variant, ByRef Index As Variant, ByRef Count As Variant)
    Dim num As Long

    Select Case VarType(Value)
        Case vbBoolean:                     WriteBoolean CBool(Value)
        Case vbByte:                        WriteByte CByte(Value)
        Case vbLong:                        WriteInt32 CLng(Value)
        Case vbSingle:                      WriteSingle CSng(Value)
        Case vbInteger:                     WriteInt16 CInt(Value)
        Case vbDouble:                      WriteDouble CDbl(Value)
        Case vbDate:                        WriteDate CDate(Value)
        Case vbCurrency:                    WriteCurrency CCur(Value)
        Case vbString:                      WriteString CStr(Value)
        Case vbDecimal
            mOutStream.WriteBlock BitConverter.GetBytes(Value), 0, 16
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_NeedIntrinsicType), "Value")
    End Select
    
    If num > 0 Then
        mBufferSA.pvData = VarPtr(Value) + VARIANTDATA_OFFSET
        If VariantType(Value) And VT_BYREF Then
            mBufferSA.pvData = MemLong(mBufferSA.pvData)
        End If
        
        mOutStream.WriteBlock mBuffer, 0, num
    End If
End Sub


Private Sub WriteArray(ByRef Value As Variant, ByRef Index As Variant, ByRef Count As Variant)
    If cArray.IsNull(Value) Then _
        Throw Cor.NewArgumentNullException("Value", Environment.GetResourceString(ArgumentNull_Array))
        
    Select Case VarType(Value)
        Case vbByteArray
            SAPtr(mBytes) = GetArrayPointer(Value)
            
            Dim ElemIndex As Long
            Dim ElemCount As Long
            Dim Result As Long
            Result = GetOptionalArrayRange(SAPtr(mBytes), Index, ElemIndex, Count, ElemCount)
            If Result <> NO_ERROR Then _
                ThrowArrayRangeException Result, "Array", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index)
            
            mOutStream.WriteBlock mBytes, ElemIndex, ElemCount
        Case vbIntegerArray
            WriteBytes mEncoding.GetBytes(Value, Index, Count)
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_UnsupportedArray), "Value")
    End Select
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    With mBufferSA
        .cbElements = 1
        .cDims = 1
        .cElements = &H10
    End With
    SAPtr(mBuffer) = VarPtr(mBufferSA)
End Sub

Private Sub Class_Terminate()
    SAPtr(mBuffer) = 0
    SAPtr(mBytes) = 0
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
