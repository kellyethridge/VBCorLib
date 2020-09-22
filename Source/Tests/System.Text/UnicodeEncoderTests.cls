VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnicodeEncoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UnicodeEncoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetByteCount_WithSingleCall_ReturnsExpected"
            Test.Use NewChars("a", "b", "c"), 6
            Test.Use NewIntegers(&HD800, &HDC00), 4
            Test.Use NewIntegers(&HD800, &HDC00, &HD800), 4
            Test.Use NewIntegers(), 0
            Test.Use NewIntegers(0), 2
            Test.Use NewIntegers(&HDC00), 2
            
        Case "GetByteCount_WithMultipleCalls_DoesNotStoreState"
            Test.Use NewIntegers(&HD800), NewIntegers(&HDC00), 2
            Test.Use NewIntegers(&H41), NewIntegers(), 2
            Test.Use NewIntegers(), NewIntegers(&H41), 2
            Test.Use NewIntegers(), NewIntegers(), 0
            Test.Use NewIntegers(&HDC00), NewIntegers(&HDC00), 4
        
        Case "GetBytes_WithSingleCall_ReturnsExpected"
            Test.Use NewChars("a", "b", "c"), NewBytes(&H61, 0, &H62, 0, &H63, 0)
            Test.Use NewIntegers(&HD800, &HDC00), NewBytes(0, &HD8, 0, &HDC)
            Test.Use NewIntegers(&HD800), NewBytes()
            Test.Use NewIntegers(&HDC00), NewBytes(&HFD, &HFF)
            
        Case "GetBytes_WithMultipleCalls_ReturnsExpected"
            Test.Use NewChars("a"), NewChars("b"), NewBytes(&H61, 0), NewBytes(&H62, 0), 4
            Test.Use NewIntegers(&HD800), NewIntegers(&HDC00), NewBytes(), NewBytes(0, &HD8, 0, &HDC), 4
            Test.Use NewIntegers(0), NewIntegers(&HD800), NewBytes(0, 0), NewBytes(&HFD, &HFF), 4
        
    End Select
End Sub

Public Sub GetByteCount_WithSingleCall_ReturnsExpected(ByRef Chars() As Integer, ByVal Expected As Long)
    Dim Encoder As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder
    
    Dim Actual As Long
    Actual = Encoder.GetByteCount(Chars, 0, CorArray.Length(Chars), False)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithMultipleCalls_DoesNotStoreState(ByRef c1() As Integer, ByRef c2() As Integer, ByVal Expected As Long)
    Dim Encoder As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder
    
    Dim Actual As Long
    Actual = Encoder.GetByteCount(c1, 0, CorArray.Length(c1), False)
    Actual = Actual + Encoder.GetByteCount(c2, 0, CorArray.Length(c2), False)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithHighSurrogateAndFlushing_ReturnsTwo()
    Dim Encoder As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder
    
    Dim Actual As Long
    Actual = Encoder.GetByteCount(NewIntegers(&HD800), 0, 1, True)
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub GetByteCount_WithHangingHighSurrogateAndFlushing_ThrowsEncoderFallbackException()
    Dim Chars() As Integer
    Dim Encoder As Encoder
    
    Chars = NewChars("a", "b", &HD800)
    Set Encoder = NewUnicodeEncoding(False, False, True).GetEncoder
    
    On Error Resume Next
    Encoder.GetByteCount Chars, 0, 3, True
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(2), "Index"
    Assert.That Ex.CharUnknown, Iz.EqualTo(&HD800), "CharUnknown"
End Sub

Public Sub GetByteCount_WithHighSurrogateAndNotFlushing_ReturnsZero()
    Dim Encoder As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder
    
    Dim Actual As Long
    Actual = Encoder.GetByteCount(NewIntegers(&HD800), 0, 1, False)
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub GetBytes_WithSingleCall_ReturnsExpected(ByRef Chars() As Integer, ByRef ExpectedBytes() As Byte)
    Dim Encoder As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder

    Dim Actual As Long
    Dim Bytes() As Byte
    Bytes = CorArray.CreateInstance(vbByte, Encoder.GetByteCount(Chars, 0, CorArray.Length(Chars), False))
    Actual = Encoder.GetBytes(Chars, 0, CorArray.Length(Chars), Bytes, 0, False)
    
    Assert.That Actual, Iz.EqualTo(CorArray.Length(ExpectedBytes))
    Assert.That Bytes, Iz.EqualTo(ExpectedBytes)
End Sub

Public Sub GetBytes_WithMultipleCalls_ReturnsExpected(ByRef c1() As Integer, ByRef c2() As Integer, ByRef ExpectedBytes1() As Byte, ByRef ExpectedBytes2() As Byte, ByVal ExpectedLength As Long)
    Dim Encoder As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder
    
    Dim ActualLength As Long
    Dim ActualBytes1() As Byte
    Dim ActualBytes2() As Byte
    ActualBytes1 = CorArray.CreateInstance(vbByte, Encoder.GetByteCount(c1, 0, CorArray.Length(c1), False))
    ActualLength = Encoder.GetBytes(c1, 0, CorArray.Length(c1), ActualBytes1, 0, False)
    ActualBytes2 = CorArray.CreateInstance(vbByte, Encoder.GetByteCount(c2, 0, CorArray.Length(c2), True))
    ActualLength = ActualLength + Encoder.GetBytes(c2, 0, CorArray.Length(c2), ActualBytes2, 0, True)
    
    Assert.That ActualLength, Iz.EqualTo(ExpectedLength)
    Assert.That ActualBytes1, Iz.EqualTo(ExpectedBytes1)
    Assert.That ActualBytes2, Iz.EqualTo(ExpectedBytes2)
End Sub

Public Sub GetBytes_WithLastCharUsed_ClearsLastChar()
    Dim Encoder As Encoder
    Dim Bytes(10) As Byte
    
    Set Encoder = NewUnicodeEncoding(False, False, True).GetEncoder
    
    Assert.That Encoder.GetBytes(NewChars(&HD800), 0, 1, Bytes, 0, False), Iz.EqualTo(0)
    Assert.That Encoder.GetBytes(NewChars(&HDC00), 0, 1, Bytes, 0, True), Iz.EqualTo(4)
    Assert.That Encoder.GetBytes(NewChars("a"), 0, 1, Bytes, 0, True), Iz.EqualTo(2)
End Sub

Public Sub Fallback_WhenCalled_ReturnsSameInstanceAsEncoding()
    Dim Encoder As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder
    
    Dim Actual As EncoderFallback
    Set Actual = Encoder.Fallback
    
    Assert.That Actual, Iz.SameAs(Encoding.Unicode.EncoderFallback)
End Sub

Public Sub Fallback_WhenSetToNothing_ThrowsArgumentNullException()
    Dim Encoder As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder
    
    On Error Resume Next
    Set Encoder.Fallback = Nothing
    
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub FallbackBuffer_WhenCalled_ReturnsCachedInstance()
    Dim Encoder As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder
    
    Dim b1 As EncoderFallbackBuffer
    Dim b2 As EncoderFallbackBuffer
    Set b1 = Encoder.FallbackBuffer
    Set b2 = Encoder.FallbackBuffer
    
    Assert.That b1, Iz.Not.Nothing
    Assert.That b1, Iz.SameAs(b2)
End Sub

Public Sub Reset_WhenCalled_ClearsInternalState()
    Dim Encoder As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder
    
    Dim b1()        As Byte
    Dim b2(0 To 1)  As Byte
    Dim Actual      As Long
    b1 = NewBytes()
    Actual = Encoder.GetBytes(NewIntegers(&HD800), 0, 1, b1, 0, False)
    Encoder.Reset
    Actual = Actual + Encoder.GetBytes(NewIntegers(&HDC00), 0, 1, b2, 0, False)
    
    Assert.That Actual, Iz.EqualTo(2)
    Assert.That b2, Iz.EqualTo(NewBytes(&HFD, &HFF))
End Sub

Public Sub Convert_WithLimitedOutputBuffer_ConvertsWithinLimits()
    Dim Chars(10000)    As Integer
    Dim Bytes(10000)    As Byte
    Dim CharsUsed       As Long
    Dim BytesUsed       As Long
    Dim Completed       As Boolean
    Dim Encoder         As Encoder
    Set Encoder = Encoding.Unicode.GetEncoder
    
    Encoder.Convert Chars, 0, 10000, Bytes, 0, 2000, True, CharsUsed, BytesUsed, Completed
    
    Assert.That CharsUsed, Iz.EqualTo(625)
    Assert.That BytesUsed, Iz.EqualTo(1250)
    Assert.That Completed, Iz.False
End Sub


