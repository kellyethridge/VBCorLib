VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryWriterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BinaryWriterTests
'
Option Explicit
Implements ITestCaseSource

Private mOutStream  As MemoryStream
Private mWriter     As BinaryWriter


Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "WriteValue_WithValues_WritesExpected"
            Test.Use CByte(123), NewBytes(123)
            Test.Use CInt(&H1234), NewBytes(&H34, &H12)
            Test.Use CLng(&H12345678), NewBytes(&H78, &H56, &H34, &H12)
            Test.Use CDbl(1.1), NewBytes(&H9A, &H99, &H99, &H99, &H99, &H99, &HF1, &H3F)
            Test.Use CSng(1.1), NewBytes(&HCD, &HCC, &H8C, &H3F)
            Test.Use CDec(1.1), NewBytes(&HB, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H1, &H0)
            Test.Use "hello", NewBytes(&H5, &H68, &H65, &H6C, &H6C, &H6F)
            Test.Use "", NewBytes(0)
            Test.Use #1/1/2001 8:30:00 AM#, NewBytes(&H55, &H55, &H55, &H55, &H8B, &H3, &HE2, &H40)
            Test.Use True, NewBytes(1)
            Test.Use False, NewBytes(0)
            Test.Use NewBytes(1, 2, 3), NewBytes(1, 2, 3)
            Test.Use NewIntegers(Asc("a"), &H2030, Asc("c")), NewBytes(&H61, &HE2, &H80, &HB0, &H63)
            
        Case "WriteChar_WithValues_WritesExpected"
            Test.Use ChrW(&H2030), NewBytes(&HE2, &H80, &HB0)
            Test.Use "Hello", NewBytes(Asc("H"))
            
        Case "WriteBoolean_WithValues_WritesExpected"
            Test.Use True, NewBytes(1)
            Test.Use False, NewBytes(0)
        
        Case "WriteBytes_WithIndexAndCount_WritesExpected"
            Test.Use 0, 10, NewBytes(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
            Test.Use 0, 5, NewBytes(1, 2, 3, 4, 5)
            Test.Use 5, 5, NewBytes(6, 7, 8, 9, 10)
            
        Case "WriteChars_WithIndexAndCount_WritesExpected"
            Test.Use 0, 7, NewBytes(65, 66, 67, 68, 69, 70, 71)
            Test.Use 0, 2, NewBytes(65, 66)
            Test.Use 4, 3, NewBytes(69, 70, 71)
            
    End Select
End Sub

Public Sub Setup()
    Set mOutStream = New MemoryStream
    Set mWriter = NewBinaryWriter(mOutStream)
End Sub

Public Sub NewBinaryWriter_WithNullStream_ThrowsArgumentNullException()
    On Error Resume Next
    
    Cor.NewBinaryWriter Nothing
    
    AssertArgumentNullException Err, "Output"
End Sub

Public Sub NewBinaryWriter_WithNonWritableStream_ThrowsArgumentException()
    Dim NonWritable As New FakeStream
    NonWritable.CanWrite = False
    
    On Error Resume Next
    Cor.NewBinaryWriter NonWritable
    
    AssertArgumentException Err
End Sub

Public Sub WriteValue_WithValues_WritesExpected(ByRef Value As Variant, ByVal Expected As Variant)
    Dim ms As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(ms)
    
    Writer.WriteValue Value
    
    Dim Actual() As Byte
    Actual = ms.ToArray
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteChar_WithValues_WritesExpected(ByRef Value As Variant, ByVal Expected As Variant)
    Dim ms As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(ms)
    
    Writer.WriteChar Value
    
    Dim Actual() As Byte
    Actual = ms.ToArray
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'Public Sub WriteValue_WithStringSubstring_WritesExpected()
'    Dim ms As New MemoryStream
'    Dim Writer As BinaryWriter
'    Set Writer = Cor.NewBinaryWriter(ms)
'
'    Writer.WriteValue "hello, world", 2, 5
'
'    Dim Actual() As Byte
'    Actual = ms.ToArray
'    Assert.That Actual, Iz.EqualTo(NewBytes(&H5, &H6C, &H6C, &H6F, &H2C, &H20))
'End Sub

Public Sub WriteValue_WithByteArray_WritesArray()
    Dim ms As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(ms)
    
    Writer.WriteValue NewBytes(1, 2, 3), 0, 3
    
    Assert.That ms.ToArray, Iz.EqualTo(NewBytes(1, 2, 3))
End Sub

Public Sub WriteValue_WithNullArray_ThrowsArgumentNullException()
    Dim NullArray() As Byte
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    
    On Error Resume Next
    Writer.WriteValue NullArray
    
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub WriteValue_WithUnsupportArrayType_ThrowsArgumentException()
    Dim UnsupportedArray(0) As Long
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    
    On Error Resume Next
    Writer.WriteValue UnsupportedArray

    AssertArgumentException Err
End Sub

Public Sub WriteBoolean_WithValues_WritesExpected(ByVal Value As Boolean, ByVal Expected As Variant)
    mWriter.WriteBoolean Value
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(Expected)
End Sub

Public Sub WriteByte_WithValue_WritesExpected()
    mWriter.WriteByte 23
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(NewBytes(23))
End Sub

Public Sub WriteInt16_WithValue_WritesExpected()
    mWriter.WriteInt16 &H1234
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(NewBytes(&H34, &H12))
End Sub

Public Sub WriteInt32_WithValue_WritesExpected()
    mWriter.WriteInt32 &H12345678
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(NewBytes(&H78, &H56, &H34, &H12))
End Sub

Public Sub WriteSingle_WithValue_WritesExpected()
    mWriter.WriteSingle 1.1
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(NewBytes(&HCD, &HCC, &H8C, &H3F))
End Sub

Public Sub WriteDouble_WithValue_WritesExpected()
    mWriter.WriteDouble 1.1
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(NewBytes(&H9A, &H99, &H99, &H99, &H99, &H99, &HF1, &H3F))
End Sub

Public Sub WriteCurrency_WithValue_WritesExpected()
    Const Value As Currency = 123456789.2222@
    Dim Expected() As Byte
    Expected = AsOByteArr(Value).Bytes
    
    mWriter.WriteCurrency Value
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(Expected)
End Sub

Public Sub WriteDecimal_WithValue_WritesExpected()
    mWriter.WriteDecimal "123456.123456"
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(NewBytes(&H40, &HF2, &H8E, &HBE, &H1C, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H6, &H0))
End Sub

Public Sub WriteData_WithValue_WritesExpected()
    mWriter.WriteDate #1/1/2001 8:30:00 AM#
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(NewBytes(&H55, &H55, &H55, &H55, &H8B, &H3, &HE2, &H40))
End Sub

Public Sub WriteBytes_WithNullArray_ThrowsArgumentNullException()
    Dim Buffer() As Byte
    
    On Error Resume Next
    mWriter.WriteBytes Buffer
    
    AssertArgumentNullException Err, "Buffer"
End Sub

Public Sub WriteBytes_WithValue_WritesExpected()
    Dim Expected() As Byte
    Expected = NewBytes(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    
    mWriter.WriteBytes Expected
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(Expected)
End Sub

Public Sub WriteBytes_WithIndexAndCount_WritesExpected(ByVal Index As Long, ByVal Count As Long, ByVal Expected As Variant)
    Dim Buffer() As Byte
    Buffer = NewBytes(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    
    mWriter.WriteBytes Buffer, Index, Count
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(Expected)
End Sub

Public Sub WriteBytes_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Buffer(5 To 10) As Byte
    
    On Error Resume Next
    mWriter.WriteBytes Buffer, 4, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub WriteBytes_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Buffer(0) As Byte
    
    On Error Resume Next
    mWriter.WriteBytes Buffer, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub WriteBytes_WithIndexCountExtendingPassedUBound_ThrowsArgumentException()
    Dim Buffer(0 To 9) As Byte
    
    On Error Resume Next
    mWriter.WriteBytes Buffer, 0, 11
    
    AssertArgumentException Err
End Sub

Public Sub WriteChars_WithNullArray_ThrowsArgumentNullException()
    Dim Chars() As Integer
    
    On Error Resume Next
    mWriter.WriteChars Chars
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub WriteChars_WithValue_WritesExpected()
    Dim Chars() As Integer
    Chars = NewIntegers(Asc("a"), &H2030, Asc("c"))

    mWriter.WriteChars Chars

    Dim Expected() As Byte
    Expected = NewBytes(&H61, &HE2, &H80, &HB0, &H63)
    Assert.That mOutStream.ToArray, Iz.EqualTo(Expected)
End Sub

Public Sub WriteChars_WithIndexAndCount_WritesExpected(ByVal Index As Long, ByVal Count As Long, ByVal Expected As Variant)
    Dim Chars() As Integer
    Chars = NewIntegers(65, 66, 67, 68, 69, 70, 71)
    
    mWriter.WriteChars Chars, Index, Count
    
    Assert.That mOutStream.ToArray, Iz.EqualTo(Expected)
End Sub

Public Sub WriteChars_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Chars(5 To 10) As Integer
    
    On Error Resume Next
    mWriter.WriteChars Chars, 4, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub WriteChars_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Chars(0) As Integer
    
    On Error Resume Next
    mWriter.WriteChars Chars, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub WriteChars_WithIndexAndCountExtendingPassedUBount_ThrowsArgumentException()
    Dim Chars(0 To 9) As Integer
    
    On Error Resume Next
    mWriter.WriteChars Chars, 0, 11
    
    AssertArgumentException Err
End Sub
