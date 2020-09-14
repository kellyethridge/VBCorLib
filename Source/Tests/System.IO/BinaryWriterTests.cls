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
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "WriteValue_WithValues_WritesExpected"
            Test.Use CByte(123), NewBytes(123)
            Test.Use CInt(&H1234), NewBytes(&H34, &H12)
            Test.Use CLng(&H12345678), NewBytes(&H78, &H56, &H34, &H12)
            Test.Use CLng(&H0), NewBytes(0, 0, 0, 0)
            Test.Use CDbl(1.1), NewBytes(&H9A, &H99, &H99, &H99, &H99, &H99, &HF1, &H3F)
            Test.Use CSng(1.1), NewBytes(&HCD, &HCC, &H8C, &H3F)
            Test.Use CDec(1.1), NewBytes(&HB, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H1, &H0)
            Test.Use "hello", NewBytes(&H5, &H68, &H65, &H6C, &H6C, &H6F)
            Test.Use "", NewBytes(0)
            Test.Use #1/1/2001 8:30:00 AM#, NewBytes(&H0, &H34, &H1A, &H2A, &HE4, &H41, &HC2, &H8)
            Test.Use True, NewBytes(1)
            Test.Use False, NewBytes(0)
            Test.Use NewBytes(1, 2, 3), NewBytes(1, 2, 3)
            Test.Use NewIntegers(Asc("a"), &H2030, Asc("c")), NewBytes(&H61, &HE2, &H80, &HB0, &H63)
            Test.Use NewInt64(&H4030201, &H8070605), NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
            
        Case "WriteChar_WithValues_WritesExpected"
            Test.Use ChrW(&H2030), NewBytes(&HE2, &H80, &HB0)
            Test.Use &H2030, NewBytes(&HE2, &H80, &HB0)
            Test.Use "Hello", NewBytes(Asc("H"))
            Test.Use CByte(49), NewBytes(49)
            Test.Use CInt(49), NewBytes(49)
            Test.Use CLng(49), NewBytes(49)
            
        Case "WriteValue_WithBytesIndexAndCount_WritesExpected"
            Test.Use 0, 10, NewBytes(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
            Test.Use 0, 5, NewBytes(1, 2, 3, 4, 5)
            Test.Use 5, 5, NewBytes(6, 7, 8, 9, 10)
            
        Case "WriteValue_WithCharsIndexAndCount_WritesExpected"
            Test.Use 0, 7, NewBytes(65, 66, 67, 68, 69, 70, 71)
            Test.Use 0, 2, NewBytes(65, 66)
            Test.Use 4, 3, NewBytes(69, 70, 71)
    End Select
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
    Dim Output As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(Output)

    Writer.WriteValue Value

    Dim Actual() As Byte
    Actual = Output.ToArray
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteChar_WithValues_WritesExpected(ByRef Value As Variant, ByVal Expected As Variant)
    Dim Output As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(Output)
    
    Writer.WriteChar Value
    
    Dim Actual() As Byte
    Actual = Output.ToArray
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteChar_WhenEmptyString_ThrowsArgumentException()
    Dim Output As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(Output)

    On Error Resume Next
    Writer.WriteChar ""
    
    AssertArgumentException Err
End Sub

Public Sub WriteChar_WithValueOutOfRange_ThrowsArgumentOutOfRangeException()
    Dim Output As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(Output)

    On Error Resume Next
    Writer.WriteChar &H10040
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub WriteValue_WithNullByteArray_ThrowsArgumentNullException()
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

Public Sub WriteValue_WithUnsupportedUDT_ThrowsArgumentException()
    Dim Unsupported As SimpleStructure
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    
    On Error Resume Next
    Writer.WriteValue Unsupported
    
    AssertArgumentException Err
End Sub

Public Sub WriteValue_WithBytes_WritesExpected()
    Dim Output As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(Output)
    Dim Bytes() As Byte
    Bytes = NewBytes(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    Writer.WriteValue Bytes
    
    Assert.That Output.ToArray, Iz.EqualTo(Bytes)
End Sub

Public Sub WriteValue_WithBytesIndexAndCount_WritesExpected(ByVal Index As Long, ByVal Count As Long, ByVal Expected As Variant)
    Dim Output As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(Output)
    Dim Buffer() As Byte
    Buffer = NewBytes(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    Writer.WriteValue Buffer, Index, Count

    Assert.That Output.ToArray, Iz.EqualTo(Expected)
End Sub

Public Sub WriteValue_WithBytesIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    Dim Buffer(5 To 10) As Byte

    On Error Resume Next
    Writer.WriteValue Buffer, 4, 1

    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub WriteValue_WithBytesCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    Dim Buffer(0) As Byte

    On Error Resume Next
    Writer.WriteValue Buffer, 0, -1

    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub WriteBytes_WithIndexCountExtendingPassedUBound_ThrowsArgumentException()
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    Dim Buffer(0 To 9) As Byte

    On Error Resume Next
    Writer.WriteValue Buffer, 0, 11

    AssertArgumentException Err
End Sub

Public Sub WriteChars_WithNullCharArray_ThrowsArgumentNullException()
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    Dim NullChars() As Integer

    On Error Resume Next
    Writer.WriteValue NullChars

    AssertArgumentNullException Err, "Value"
End Sub

Public Sub WriteChars_WithCharArray_WritesExpected()
    Dim Output As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(Output)
    Dim Chars() As Integer
    Chars = NewIntegers(Asc("a"), &H2030, Asc("c"))

    Writer.WriteValue Chars

    Dim Expected() As Byte
    Expected = NewBytes(&H61, &HE2, &H80, &HB0, &H63)
    Assert.That Output.ToArray, Iz.EqualTo(Expected)
End Sub

Public Sub WriteValue_WithCharsIndexAndCount_WritesExpected(ByVal Index As Long, ByVal Count As Long, ByVal Expected As Variant)
    Dim Output As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(Output)
    Dim Chars() As Integer
    Chars = NewIntegers(65, 66, 67, 68, 69, 70, 71)

    Writer.WriteValue Chars, Index, Count

    Assert.That Output.ToArray, Iz.EqualTo(Expected)
End Sub

Public Sub WriteChars_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    Dim Chars(5 To 10) As Integer

    On Error Resume Next
    Writer.WriteValue Chars, 4, 1

    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub WriteChars_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    Dim Chars(0) As Integer

    On Error Resume Next
    Writer.WriteValue Chars, 0, -1

    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub WriteChars_WithIndexAndCountExtendingPassedUBount_ThrowsArgumentException()
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    Dim Chars(0 To 9) As Integer

    On Error Resume Next
    Writer.WriteValue Chars, 0, 11

    AssertArgumentException Err
End Sub

Public Sub SeekPosition_WhenCalled_ReturnsNewPosition()
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(New MemoryStream)
    Writer.WriteValue NewBytes(1, 2, 3, 4, 5)
    
    Dim Actual As Long
    Actual = Writer.SeekPosition(1, FromEnd)
    
    Assert.That Actual, Iz.EqualTo(6)
End Sub
