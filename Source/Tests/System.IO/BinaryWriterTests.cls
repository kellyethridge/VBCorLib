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

Public Sub WriteValue_WithLargeString_WritesString()
    Dim Output As New MemoryStream
    Dim Writer As BinaryWriter
    Dim s As String
    Dim i As Long
    
    Set Writer = Cor.NewBinaryWriter(Output)
    
    For i = 1 To 300
        s = s & ChrW$(i)
    Next
    
    Writer.WriteValue s
    
    Dim Expected() As Byte
    Expected = NewBytes(&HD9, &H3, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F, &H20, &H21, &H22, &H23, &H24, &H25, &H26, &H27, &H28, &H29, &H2A, &H2B, &H2C, &H2D, &H2E, &H2F, &H30, &H31, &H32, &H33, &H34, &H35, &H36, &H37, &H38, &H39, &H3A, &H3B, &H3C, &H3D, &H3E, &H3F, &H40, &H41, &H42, &H43, &H44, &H45, &H46, &H47, &H48, &H49, &H4A, &H4B, &H4C, &H4D, &H4E, &H4F, &H50, &H51, &H52, &H53, _
                        &H54, &H55, &H56, &H57, &H58, &H59, &H5A, &H5B, &H5C, &H5D, &H5E, &H5F, &H60, &H61, &H62, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &H6A, &H6B, &H6C, &H6D, &H6E, &H6F, &H70, &H71, &H72, &H73, &H74, &H75, &H76, &H77, &H78, &H79, &H7A, &H7B, &H7C, &H7D, &H7E, &H7F, &HC2, &H80, &HC2, &H81, &HC2, &H82, &HC2, &H83, &HC2, &H84, &HC2, &H85, &HC2, &H86, &HC2, &H87, &HC2, &H88, &HC2, &H89, &HC2, &H8A, &HC2, &H8B, &HC2, &H8C, &HC2, &H8D, &HC2, &H8E, &HC2, &H8F, &HC2, &H90, &HC2, &H91, &HC2, &H92, &HC2, _
                        &H93, &HC2, &H94, &HC2, &H95, &HC2, &H96, &HC2, &H97, &HC2, &H98, &HC2, &H99, &HC2, &H9A, &HC2, &H9B, &HC2, &H9C, &HC2, &H9D, &HC2, &H9E, &HC2, &H9F, &HC2, &HA0, &HC2, &HA1, &HC2, &HA2, &HC2, &HA3, &HC2, &HA4, &HC2, &HA5, &HC2, &HA6, &HC2, &HA7, &HC2, &HA8, &HC2, &HA9, &HC2, &HAA, &HC2, &HAB, &HC2, &HAC, &HC2, &HAD, &HC2, &HAE, &HC2, &HAF, &HC2, &HB0, &HC2, &HB1, &HC2, &HB2, &HC2, &HB3, &HC2, &HB4, &HC2, &HB5, &HC2, &HB6, &HC2, &HB7, &HC2, &HB8, &HC2, &HB9, &HC2, &HBA, &HC2, &HBB, &HC2, &HBC, _
                        &HC2, &HBD, &HC2, &HBE, &HC2, &HBF, &HC3, &H80, &HC3, &H81, &HC3, &H82, &HC3, &H83, &HC3, &H84, &HC3, &H85, &HC3, &H86, &HC3, &H87, &HC3, &H88, &HC3, &H89, &HC3, &H8A, &HC3, &H8B, &HC3, &H8C, &HC3, &H8D, &HC3, &H8E, &HC3, &H8F, &HC3, &H90, &HC3, &H91, &HC3, &H92, &HC3, &H93, &HC3, &H94, &HC3, &H95, &HC3, &H96, &HC3, &H97, &HC3, &H98, &HC3, &H99, &HC3, &H9A, &HC3, &H9B, &HC3, &H9C, &HC3, &H9D, &HC3, &H9E, &HC3, &H9F, &HC3, &HA0, &HC3, &HA1, &HC3, &HA2, &HC3, &HA3, &HC3, &HA4, &HC3, &HA5, &HC3, _
                        &HA6, &HC3, &HA7, &HC3, &HA8, &HC3, &HA9, &HC3, &HAA, &HC3, &HAB, &HC3, &HAC, &HC3, &HAD, &HC3, &HAE, &HC3, &HAF, &HC3, &HB0, &HC3, &HB1, &HC3, &HB2, &HC3, &HB3, &HC3, &HB4, &HC3, &HB5, &HC3, &HB6, &HC3, &HB7, &HC3, &HB8, &HC3, &HB9, &HC3, &HBA, &HC3, &HBB, &HC3, &HBC, &HC3, &HBD, &HC3, &HBE, &HC3, &HBF, &HC4, &H80, &HC4, &H81, &HC4, &H82, &HC4, &H83, &HC4, &H84, &HC4, &H85, &HC4, &H86, &HC4, &H87, &HC4, &H88, &HC4, &H89, &HC4, &H8A, &HC4, &H8B, &HC4, &H8C, &HC4, &H8D, &HC4, &H8E, &HC4, &H8F, _
                        &HC4, &H90, &HC4, &H91, &HC4, &H92, &HC4, &H93, &HC4, &H94, &HC4, &H95, &HC4, &H96, &HC4, &H97, &HC4, &H98, &HC4, &H99, &HC4, &H9A, &HC4, &H9B, &HC4, &H9C, &HC4, &H9D, &HC4, &H9E, &HC4, &H9F, &HC4, &HA0, &HC4, &HA1, &HC4, &HA2, &HC4, &HA3, &HC4, &HA4, &HC4, &HA5, &HC4, &HA6, &HC4, &HA7, &HC4, &HA8, &HC4, &HA9, &HC4, &HAA, &HC4, &HAB, &HC4, &HAC)
    Assert.That Output.ToArray, Iz.EqualTo(Expected)
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

Public Sub NewBinaryWriter_WithLeaveOpenTrue_DoesNotCloseUnderlyingStream()
    Dim ms As New MemoryStream
    Dim Writer As BinaryWriter
    
    Set Writer = Cor.NewBinaryWriter(ms, LeaveOpen:=True)
    Writer.WriteValue CByte(1)
    Writer.CloseWriter
    
    ms.WriteByte 2
    
    Assert.That ms.Position, Iz.EqualTo(2)
End Sub
