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
Option Explicit

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

Public Sub ReadLong_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HFF, &H0, &H80, &H80))
    
    Dim Actual As Long
    Actual = Reader.ReadLong
    
    Assert.That Actual, Iz.EqualTo(&H808000FF)
End Sub

Public Sub ReadInteger_WhenCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H98, &H3A))
    
    Dim Actual As Integer
    Actual = Reader.ReadInteger
    
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
    Set Reader = MakeReader(NewBytes(&HAB, &HAA, &HAA, &HAA, &HEA, &H5E, &HE2, &H40))
    
    Dim Actual As Date
    Actual = Reader.ReadDate
    
    Assert.That Actual, Iz.EqualTo(#1/2/2003 8:00:00 AM#)
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

Public Sub ReadXxx_WithMultipleCalled_ReturnsExpected()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&H44, &H0, &H0, &H0, &H8, &H68, &H65, &H6C, &H6C, &H6F, &HE2, &H80, &HB0, &HCD, &HCC, &HCC, &HCC, &HCC, &H1C, &H51, &H40))
    
    Assert.That Reader.ReadLong, Iz.EqualTo(68)
    Assert.That Reader.ReadString, Iz.EqualTo("hello" & ChrW$(&H2030))
    Assert.That Reader.ReadDouble, Iz.EqualTo(68.45)
End Sub

Public Sub ReadChar_WhenCalled_ReturnsChar()
    Dim Reader As BinaryReader
    Set Reader = MakeReader(NewBytes(&HE2, &H80, &HB0))
    
    Dim Actual As Long
    Actual = Reader.ReadChar
    
    Assert.That Actual, Iz.EqualTo(&H2030)
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

Private Function MakeReader(ByRef Data() As Byte) As BinaryReader
    Dim Source As New MemoryStream
    Source.WriteBlock Data, 0, UBound(Data) + 1
    Source.Position = 0
    Set MakeReader = Cor.NewBinaryReader(Source)
End Function
