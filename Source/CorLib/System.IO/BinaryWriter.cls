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
' @remarks The data types that are written to the stream are written in their binary
' form. The values are written as they are represented in memory. For example, as data type
' of vbLong is written as 4 bytes, regardless of the actual value of the vbLong variable.
'
' @see Constructors
' @see BinaryReader
Option Explicit
Implements IObject

Private Const LargeBufferSize As Long = 256

Private mOutStream      As Stream
Private mEncoding       As Encoding
Private mSingleChar(0)  As Integer
Private mCharBytes(5)   As Byte
Private mLargeBuffer()  As Byte
Private mMaxChars       As Long
Private mLeaveOpen      As Boolean


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
    If Not mOutStream Is Nothing Then
        If Not mLeaveOpen Then
            mOutStream.CloseStream
        End If
        
        Set mOutStream = Nothing
    End If
End Sub

''
' Tells the underlying stream to store its data to its repository.
'
Public Sub Flush()
    VerifyIsOpen
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
    VerifyIsOpen
    SeekPosition = mOutStream.SeekPosition(Offset, Origin)
End Function

''
' Writes the binary representation of a datatype to the underlying stream.
'
' @param value The datatype to be written.
' @param Index The starting index of arrays or string characters to being writing from.
' @param Count The number of elements or characters to be written to the stream.
' @remarks <i>Index</i> and <i>vCount</i> are only valid when writing Byte arrays,
' Integer array, and Strings. For all other types, these parameters are ignored.
' <p>Simple datatypes such as vbLong and vbDouble are written out exactly as they
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
        ValidateArrayRange Value, Index, Count, NameOfValue
        WriteArray Value, Index, Count
    Else
        WriteIntrinsic Value
    End If
End Sub

''
' Writes a character to the underlying stream.
'
' @param Value The character to be written. This can be a string or a numeric value.
' @remarks The character is encoded and the byte array is written to the stream.
'
Public Sub WriteChar(ByRef Value As Variant)
    Dim BytesRead As Long
    
    VerifyIsOpen
    
    Select Case VarType(Value)
        Case vbString
            If Len(Value) = 0 Then _
                Error.Argument Argument_StringZeroLength
            
            mSingleChar(0) = AscW(CStr(Value))
        Case vbLong, vbInteger, vbByte
            mSingleChar(0) = Char.ConvertFromInt32(Value)
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidValueType, "String or Integer"), "Value")
    End Select
    
    BytesRead = mEncoding.GetBytesEx(mSingleChar, 0, 1, mCharBytes, 0)
    
    If BytesRead > 0 Then
        mOutStream.WriteBlock mCharBytes, 0, BytesRead
    End If
End Sub

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
' @param value The value to compare equalit to.
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
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Output As Stream, ByVal Encoding As Encoding, ByVal LeaveOpen As Boolean)
    If Output Is Nothing Then _
        Error.ArgumentNull "Output", ArgumentNull_Stream
    If Not Output.CanWrite Then _
        Error.Argument Argument_StreamNotWritable
    
    Set mOutStream = Output
    mLeaveOpen = LeaveOpen
    
    If Encoding Is Nothing Then
        Set mEncoding = Cor.NewUTF8Encoding(False, True)
    Else
        Set mEncoding = Encoding
    End If
    
    ReDim mLargeBuffer(0 To LargeBufferSize - 1)
    mMaxChars = LargeBufferSize \ mEncoding.GetMaxByteCount(1)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyIsOpen()
    If mOutStream Is Nothing Then
        Error.ObjectDisposed ObjectDisposed_FileNotOpen
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

Private Sub WriteIntrinsic(ByRef Value As Variant)
    Select Case VarType(Value)
        Case vbBoolean
            WriteBoolean Value
        Case vbByte
            WriteByte Value
        Case vbLong
            WriteInt32 Value
        Case vbSingle
            WriteSingle Value
        Case vbInteger
            WriteInt16 Value
        Case vbDouble
            WriteDouble Value
        Case vbDate
            WriteDate Value
        Case vbCurrency
            WriteCurrency Value
        Case vbString
            WriteString Value
        Case vbDecimal
            WriteDecimal Value
        Case Else
            If IsInt64(Value) Then
                WriteInt64 Value
            Else
                Error.Argument Argument_NeedIntrinsicType, "Value"
            End If
    End Select
End Sub

Private Sub WriteInt64(ByRef Value As Variant)
    mOutStream.WriteBlock AsOByteArr(ByVal DataPtr(Value)).Bytes, 0, vbSizeOfInt64
End Sub

Private Sub WriteBoolean(ByVal Value As Boolean)
    If Value Then
        mOutStream.WriteByte 1
    Else
        mOutStream.WriteByte 0
    End If
End Sub

Private Sub WriteByte(ByVal Value As Byte)
    mOutStream.WriteByte Value
End Sub

Private Sub WriteInt32(ByVal Value As Long)
    mOutStream.WriteBlock AsQByteArr(Value).Bytes, 0, vbSizeOfLong
End Sub

Private Sub WriteInt16(ByVal Value As Integer)
    mOutStream.WriteBlock AsDByteArr(Value).Bytes, 0, vbSizeOfInteger
End Sub

Private Sub WriteSingle(ByVal Value As Single)
    mOutStream.WriteBlock AsQByteArr(Value).Bytes, 0, vbSizeOfSingle
End Sub

Private Sub WriteDouble(ByVal Value As Double)
    mOutStream.WriteBlock AsOByteArr(Value).Bytes, 0, vbSizeOfDouble
End Sub

Private Sub WriteDate(ByVal Value As Date)
    Dim Milliseconds As Currency
    
    Milliseconds = CorDateTime.GetTotalMilliseconds(Value)
    WriteCurrency Milliseconds
End Sub

Private Sub WriteCurrency(ByVal Value As Currency)
    mOutStream.WriteBlock AsOByteArr(Value).Bytes, 0, vbSizeOfCurrency
End Sub

Private Sub WriteDecimal(ByRef Value As Variant)
    mOutStream.WriteBlock BitConverter.GetBytes(Value), 0, vbSizeOfDecimal
End Sub

' Keep the string in the Variant type since all our methods will
' accept a variant type, thus preventing string re-allocation.
Private Sub WriteString(ByRef Value As Variant)
    WriteStringLength mEncoding.GetByteCount(Value)
    
    If Len(Value) > mMaxChars Then
        ' This is a larger than expected string so we will
        ' go the old route and let a buffer be created and destroyed.
        Dim Bytes() As Byte

        Bytes = mEncoding.GetBytes(Value)
        mOutStream.WriteBlock Bytes, 0, Len1D(Bytes)
    Else
        ' Most strings aren't going to be very long, so we try to
        ' re-use an existing buffer so not to constantly reallocate one.
        Dim ByteCount As Long

        ByteCount = mEncoding.GetBytesEx(Value, 0, Len(Value), mLargeBuffer, 0)
        mOutStream.WriteBlock mLargeBuffer, 0, ByteCount
    End If
End Sub

Private Sub WriteArray(ByRef Value As Variant, ByRef Index As Variant, ByRef Count As Variant)
    Select Case VarType(Value)
        Case vbByteArray
            WriteBytes Value, Index, Count
        Case vbIntegerArray
            WriteIntegers Value, Index, Count
        Case Else
            Error.Argument Argument_UnsupportedArray, "Value"
    End Select
End Sub

Private Sub WriteBytes(ByRef Buffer As Variant, ByRef Index As Variant, ByRef Count As Variant)
    Dim Bytes() As Byte
    Dim Range   As ListRange
    
    Range = MakeArrayRange(Buffer, Index, Count)
    SAPtr(Bytes) = SAPtrV(Buffer)
    
    On Error GoTo Catch
    mOutStream.WriteBlock Bytes, Range.Index, Range.Count
    
    SAPtr(Bytes) = vbNullPtr
    Exit Sub
    
Catch:
    SAPtr(Bytes) = vbNullPtr
    ThrowOrErr Err
End Sub

Private Sub WriteIntegers(ByRef Value As Variant, ByRef Index As Variant, ByRef Count As Variant)
    If Len1D(Value) > mMaxChars Then
        Dim Bytes() As Byte
        
        Bytes = mEncoding.GetBytes(Value, Index, Count)
        mOutStream.WriteBlock Bytes, 0, Len1D(Bytes)
    Else
        Dim ByteCount   As Long
        Dim Range       As ListRange
        
        Range = MakeArrayRange(Value, Index, Count)
        ByteCount = mEncoding.GetBytesEx(Value, Range.Index, Range.Count, mLargeBuffer, 0)
        mOutStream.WriteBlock mLargeBuffer, 0, ByteCount
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
