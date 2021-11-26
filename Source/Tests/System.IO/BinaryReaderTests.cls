VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryReaderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BinaryReaderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub CloseReader_WhenCalled_ClosesBaseStream()
    Dim ms As New MemoryStream
    Dim Reader As BinaryReader
    Set Reader = NewBinaryReader(ms)
    
    On Error Resume Next
    Reader.CloseReader
    ms.WriteByte 1
    
    AssertObjectDisposedException Err
End Sub

Public Sub CloseReader_WithLeaveOpenOption_DoesNotCloseBaseStream()
    Dim ms As New MemoryStream
    Dim Reader As BinaryReader
    Set Reader = NewBinaryReader(ms, LeaveOpen:=True)
    
    Reader.CloseReader
    ms.WriteByte 1

    Assert.That ms.ToArray(), Iz.EqualTo(NewBytes(1))
End Sub

Public Sub NewBinaryReader_WithNothingStream_ThrowsArgumentNullException()
    On Error Resume Next
    
    Cor.NewBinaryReader Nothing
    
    AssertArgumentNullException Err, "Stream"
End Sub

Public Sub NewBinaryReader_WithNonReadableStream_ThrowsArgumentException()
    Dim Nonreadable As New FakeStream
    Nonreadable.CanRead = False
    
    On Error Resume Next
    Cor.NewBinaryReader Nonreadable
    
    AssertArgumentException Err
End Sub

Public Sub ReadByte_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(129))
    
    Dim Actual As Long
    Actual = Reader.ReadByte
    
    Assert.That Actual, Iz.EqualTo(129)
End Sub

Public Sub ReadByte_WithMultipleCalls_ReturnsBytes()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(1, 2, 3, 4))
    
    Assert.That Reader.ReadByte, Iz.EqualTo(1)
    Assert.That Reader.ReadByte, Iz.EqualTo(2)
    Assert.That Reader.ReadByte, Iz.EqualTo(3)
    Assert.That Reader.ReadByte, Iz.EqualTo(4)
End Sub

Public Sub ReadByte_WhenReadPastEndOfStream_ThrowsEndOfStreamException()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(129))
    Reader.ReadByte
    
    On Error Resume Next
    Reader.ReadByte
    
    AssertEndOfStreamException Err
End Sub

Public Sub ReadBytes_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(129))
    
    On Error Resume Next
    Reader.ReadBytes -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub ReadBytes_WithNoBytesRemainingInStream_ReturnsZeroLengthArray()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes())
    
    Dim Actual() As Byte
    Actual = Reader.ReadBytes(1)
    
    Assert.That Actual, Iz.EqualTo(NewBytes())
End Sub

Public Sub ReadInt64_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(1, 2, 3, 4, 5, 6, 7, 8))
    
    Dim Actual As Int64
    Actual = Reader.ReadInt64
    
    Assert.That Actual, Iz.EqualTo(NewInt64(&H4030201, &H8070605))
End Sub

Public Sub ReadInt32_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HFF, &H0, &H80, &H80))
    
    Dim Actual As Long
    Actual = Reader.ReadInt32
    
    Assert.That Actual, Iz.EqualTo(&H808000FF)
End Sub

Public Sub ReadInt32_WhenNotEnoughData_ThrowsEndOfStreamException()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H1, &H2))
    
    On Error Resume Next
    Reader.ReadInt32
    
    AssertEndOfStreamException Err
End Sub

Public Sub ReadInt16_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H98, &H3A))
    
    Dim Actual As Integer
    Actual = Reader.ReadInt16
    
    Assert.That Actual, Iz.EqualTo(15000)
End Sub

Public Sub ReadDouble_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HAD, &HFA, &H5C, &H6D, &H45, &H4A, &H93, &H40))
    
    Dim Actual As Double
    Actual = Reader.ReadDouble
    
    Assert.That Actual, Iz.EqualTo(1234.5678)
End Sub

Public Sub ReadSingle_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H25, &H52, &H9A, &H44))
    
    Dim Actual As Single
    Actual = Reader.ReadSingle
    
    Assert.That Actual, Iz.EqualTo(1234.567!)
End Sub

Public Sub ReadDate_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H90, &HE7, &HD3, &HCC, &H7F, &H42, &HC2, &H8))
    
    Dim Actual As Date
    Actual = Reader.ReadDate
    
    Assert.That Actual, Iz.EqualTo(#1/2/2001 3:04:05 AM#)
End Sub

Public Sub ReadDate_WithZeroMilliseconds_ReturnsDefaultDate()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(0, 0, 0, 0, 0, 0, 0, 0))
    
    Dim Actual As Date
    Actual = Reader.ReadDate
    
    Assert.That Actual, Iz.EqualTo(#12:00:00 AM#)
End Sub

Public Sub ReadCurrency_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H0, &H40, &H2F, &H15, &H4C, &H80, &HC4, &H8))
    
    Dim Actual As Currency
    Actual = Reader.ReadCurrency
    
    Assert.That Actual, Iz.EqualTo(63177091200000@)
End Sub

Public Sub ReadDecimal_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H7, &H34, &H59, &H73, &H3A, &HB, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H7, &H0))
    
    Dim Actual As Variant
    Actual = Reader.ReadDecimal
    
    Assert.That Actual, Iz.EqualTo(CDec("1234567.1234567"))
End Sub

Public Sub ReadString_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H8, &H68, &H65, &H6C, &H6C, &H6F, &HE2, &H80, &HB0))
    
    Dim Actual As String
    Actual = Reader.ReadString
    
    Assert.That Actual, Iz.EqualTo("hello" & ChrW$(&H2030))
End Sub

Public Sub ReadString_WithInvalid7BitEncodedLength_ThrowsFormatException()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H81, &H81, &H81, &H81, &H81, 0))
    
    On Error GoTo Catch
    Reader.ReadString
    
Catch:
    AssertFormatException Err
End Sub

Public Sub ReadXxx_WithMultipleCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H44, &H0, &H0, &H0, &H8, &H68, &H65, &H6C, &H6C, &H6F, &HE2, &H80, &HB0, &HCD, &HCC, &HCC, &HCC, &HCC, &H1C, &H51, &H40))
    
    Assert.That Reader.ReadInt32, Iz.EqualTo(68)
    Assert.That Reader.ReadString, Iz.EqualTo("hello" & ChrW$(&H2030))
    Assert.That Reader.ReadDouble, Iz.EqualTo(68.45)
End Sub

Public Sub ReadChar_WhenCalled_ReturnsChar()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HE2, &H80, &HB0))
    
    Dim Actual As Integer
    Actual = Reader.ReadChar
    
    Assert.That Actual, Iz.EqualTo(&H2030)
End Sub

Public Sub ReadChar_WhenReadingCharacterWithHighBit_ReturnsConvertedCorrectly()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HEA, &H80, &H80))
    
    Dim Actual As Integer
    Actual = Reader.ReadChar
    
    Assert.That Actual, Iz.EqualTo(&HA000)
End Sub

Public Sub ReadChar_WhenReadingASurrogateFromAPair_ThrowsArgumentException()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(240, 144, 128, 128)) ' UTF-8 U+D800 - U+DC00
    
    On Error Resume Next
    Reader.ReadChar
    
    AssertArgumentException Err
End Sub

Public Sub ReadChar_WhenReadingASurrogateFromAPair_RestoresPositionInUnderlyingStreamToBeforeReadAttempt()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(240, 144, 128, 128)) ' UTF-8 U+D800 - U+DC00

    On Error GoTo Catch
    Reader.ReadChar
    
Catch:
    Dim Actual As Long
    Actual = Reader.BaseStream.Position
    
    Assert.That Actual, Iz.EqualTo(0)
    Assert.Pass
End Sub

Public Sub ReadChars_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HE2, &H80, &HB0))
    
    On Error Resume Next
    Reader.ReadChars -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub ReadChars_WhenCalled_ReturnsArrayOfChars()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HE2, &H80, &HB0, &H65))

    Dim Actual() As Integer
    Actual = Reader.ReadChars(2)
    
    Assert.That Actual, Iz.EqualTo(NewIntegers(&H2030, &H65))
End Sub

Public Sub ReadChars_WithLargeAmountUTF8_ReturnsChars()
    Const Count As Long = 1024
    Dim Reader As BinaryReader
    Dim Bytes() As Byte
    Dim Expected() As Integer
    Dim i As Long
    
    ReDim Bytes(0 To Count - 1)
    ReDim Expected(0 To Count - 1)
    
    For i = 0 To Count - 1
        Expected(i) = i
    Next
    
    Bytes = Encoding.UTF8.GetBytes(Expected)
    Set Reader = MakeReader(Bytes)
    
    Dim Actual() As Integer
    Actual = Reader.ReadChars(Count)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadChars_WhenMoreCharsRequestedThanAvailable_ReturnsAvailableChars()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HE2, &H80, &HB0, &H65))

    Dim Actual() As Integer
    Actual = Reader.ReadChars(3)
    
    Assert.That Actual, Iz.EqualTo(NewIntegers(&H2030, &H65))
End Sub

Public Sub ReadChars_WithNoCharsRemaining_ReturnsZeroLengthArray()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes)
    
    Dim Actual() As Integer
    Actual = Reader.ReadChars(1)
    
    Assert.That Actual, Iz.EqualTo(NewIntegers())
End Sub

Public Sub PeekChar_WhenCalled_ReturnsNextChar()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HE2, &H80, &HB0))
    
    Dim Actual As Long
    Actual = Reader.PeekChar
    
    Assert.That Actual, Iz.EqualTo(&H2030)
End Sub

Public Sub PeakChar_WhenCalled_DoesNotRemoveNextChar()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HE2, &H80, &HB0))
    
    Reader.PeekChar
    
    Assert.That Reader.ReadChar, Iz.EqualTo(&H2030)
End Sub

Public Sub Read_WithByteArray_FillsArray()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(1, 2, 3, 4))
    Dim Target(3) As Byte
    
    Dim Actual As Long
    Actual = Reader.Read(Target, 0, 4)
    
    Assert.That Actual, Iz.EqualTo(4)
    Assert.That Target, Iz.EqualTo(NewBytes(1, 2, 3, 4))
End Sub

Public Sub Read_WithCharacterArray_FillsArray()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H61, &H62))
    Dim Target(3) As Integer
    
    Dim Actual As Long
    Actual = Reader.Read(Target, 1, 2)
    
    Assert.That Actual, Iz.EqualTo(2)
    Assert.That Target, Iz.EqualTo(NewBytes(0, Asc("a"), Asc("b"), 0))
End Sub

Public Sub Read_WhenCalled_ReturnsNextCharacter()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HE2, &H80, &HB0))
    
    Dim Actual As Long
    Actual = Reader.Read
    
    Assert.That Actual, Iz.EqualTo(&H2030)
End Sub

Public Sub Read_WhenNoMoreCharacters_ReturnsMinusOne()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HE2, &H80, &HB0))
    Reader.Read
    
    Dim Actual As Long
    Actual = Reader.Read
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub Read_WithNullArray_ThrowsArgumentNullException()
    Dim NullArray() As Byte
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HE2, &H80, &HB0))
    
    On Error Resume Next
    Reader.Read NullArray, 0, 1
    
    AssertArgumentNullException Err, "Buffer"
End Sub

Public Sub Read_WithUnsupportedArray_ThrowsArgumentException()
    Dim Longs(0) As Long
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(0))
    
    On Error Resume Next
    Reader.Read Longs, 0, 1
    
    AssertArgumentException Err, "Buffer"
End Sub

Public Sub Read_WithMissingCount_ThrowsArgumentException()
    Dim Target(0) As Byte
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(1))
    
    On Error Resume Next
    Reader.Read Target, 0
    
    AssertArgumentException Err, "Count"
End Sub

Public Sub Read_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Target(1 To 3) As Byte
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(1, 2, 3))
    
    On Error Resume Next
    Reader.Read Target, 0, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Read_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Target(1 To 3) As Byte
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(1, 2, 3))
    
    On Error Resume Next
    Reader.Read Target, 1, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Read_WithIndexAndCountNotValidPosition_ThrowsArgumentException()
    Dim Target(1 To 3) As Byte
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(1, 2, 3))
    
    On Error Resume Next
    Reader.Read Target, 1, 4
    
    AssertArgumentException Err
End Sub

Public Sub ReadChar_WithClosedReader_ThrowsObjectDisposedException()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(1))
    Reader.CloseReader
    
    On Error Resume Next
    Reader.ReadChar
    
    AssertObjectDisposedException Err
End Sub

Public Sub ReadChar_WithNoMoreCharacters_ThrowsEndOfStreamException()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(65))
    
    Reader.ReadChar
    
    On Error Resume Next
    Reader.ReadChar
    
    AssertEndOfStreamException Err
End Sub

Private Function MakeReader(ByRef Data() As Byte) As BinaryReader
    Dim Source As New MemoryStream
    Source.WriteBlock Data, 0, UBound(Data) + 1
    Source.Position = 0
    Set MakeReader = Cor.NewBinaryReader(Source)
End Function
