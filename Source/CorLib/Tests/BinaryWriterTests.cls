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
            
    End Select
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

Public Sub WriteValue_WithStringSubstring_WritesExpected()
    Dim ms As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(ms)
    
    Writer.WriteValue "hello, world", 2, 5
    
    Dim Actual() As Byte
    Actual = ms.ToArray
    Assert.That Actual, Iz.EqualTo(NewBytes(&H5, &H6C, &H6C, &H6F, &H2C, &H20))
End Sub

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
