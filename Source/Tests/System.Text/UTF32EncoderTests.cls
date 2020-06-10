VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF32EncoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UTF32EncoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetByteCount_WithSingleCallWithOutFlushing_ReturnsExpected"
            Test.Use NewChars("a", "b", "c"), 12
            Test.Use NewChars(&HD800, &HDC00), 4
            Test.Use NewChars(&HD800, &HDC00, &HD800), 4
            Test.Use NewChars(), 0
            Test.Use NewChars(0), 4
            Test.Use NewChars(&HDC00), 4
            Test.Use NewChars(&HD800), 0
            
        Case "GetBytes_WithSingleCall_ReturnsExpected"
            Test.Use NewChars("a", "b", "c"), NewBytes(&H61, 0, 0, 0, &H62, 0, 0, 0, &H63, 0, 0, 0)
            Test.Use NewChars("a", &HD8FF, &HDC00), NewBytes(97, 0, 0, 0, 0, 252, 4, 0)
            Test.Use NewChars(&HD800), NewBytes()
            Test.Use NewChars(&HDC00), NewBytes(&HFD, &HFF, 0, 0)
            
        Case "GetBytes_WithMultipleCalls_ReturnsExpected"
            Test.Use NewChars("a"), NewChars("b"), NewBytes(&H61, 0, 0, 0), NewBytes(&H62, 0, 0, 0), 8
            Test.Use NewChars(&HD8FF), NewChars(&HDC00), NewBytes(), NewBytes(0, 252, 4, 0), 4
            Test.Use NewChars(0), NewChars(&HD800), NewBytes(0, 0, 0, 0), NewBytes(&HFD, &HFF, 0, 0), 8
            
    End Select
End Sub

Public Sub GetByteCount_WithSingleCallWithOutFlushing_ReturnsExpected(ByRef Chars() As Integer, ByVal Expected As Long)
    Dim Actual As Long
    Dim Encoder As Encoder
    
    Set Encoder = Encoding.UTF32.GetEncoder
    
    Actual = Encoder.GetByteCount(Chars, 0, CorArray.Length(Chars), False)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithTrailingHighSurrogateAndFlushing_CountsFallbackCharacters()
    Dim Encoder As Encoder
    Dim Actual As Long
    
    Set Encoder = MakeEncoder
    
    Actual = Encoder.GetByteCount(NewChars("a", &HD800), 0, 2, True)
    
    Assert.That Actual, Iz.EqualTo(12)
End Sub

Public Sub GetByteCount_WithTrailingHighSurrogateAndNotFlushing_IgnoresTrailingHighSurrogate()
    Dim Encoder As Encoder
    Dim Actual As Long
    
    Set Encoder = MakeEncoder
    
    Actual = Encoder.GetByteCount(NewChars("a", &HD800), 0, 2, False)
    
    Assert.That Actual, Iz.EqualTo(4)
End Sub

Public Sub GetByteCount_WithTrailingHighSurrogate_DoesNotStoreHighSurrogate()
    Dim Encoder As Encoder
    Dim Actual As Long
    
    Set Encoder = MakeEncoder
    
    Actual = Encoder.GetByteCount(NewChars("a", &HD800), 0, 2, False)
    Assert.That Actual, Iz.EqualTo(4)
    Actual = Encoder.GetByteCount(NewChars(&HDC00), 0, 1, False)
    Assert.That Actual, Iz.EqualTo(8)
End Sub

Public Sub GetBytes_WithSingleCall_ReturnsExpected(ByRef Chars() As Integer, ByRef ExpectedBytes() As Byte)
    Dim Encoder As Encoder
    Dim Actual As Long
    Dim Bytes() As Byte
    
    Set Encoder = Encoding.UTF32.GetEncoder

    Bytes = CorArray.CreateInstance(vbByte, Encoder.GetByteCount(Chars, 0, CorArray.Length(Chars), False))
    Actual = Encoder.GetBytes(Chars, 0, CorArray.Length(Chars), Bytes, 0, False)
    
    Assert.That Actual, Iz.EqualTo(CorArray.Length(ExpectedBytes))
    Assert.That Bytes, Iz.EqualTo(ExpectedBytes)
End Sub

Public Sub GetBytes_WithMultipleCalls_ReturnsExpected(ByRef c1() As Integer, ByRef c2() As Integer, ByRef ExpectedBytes1() As Byte, ByRef ExpectedBytes2() As Byte, ByVal ExpectedLength As Long)
    Dim Encoder As Encoder
    Dim ActualLength As Long
    Dim ActualBytes1() As Byte
    Dim ActualBytes2() As Byte
    
    Set Encoder = Encoding.UTF32.GetEncoder
    
    ActualBytes1 = CorArray.CreateInstance(vbByte, Encoder.GetByteCount(c1, 0, CorArray.Length(c1), False))
    ActualLength = Encoder.GetBytes(c1, 0, CorArray.Length(c1), ActualBytes1, 0, False)
    ActualBytes2 = CorArray.CreateInstance(vbByte, Encoder.GetByteCount(c2, 0, CorArray.Length(c2), True))
    ActualLength = ActualLength + Encoder.GetBytes(c2, 0, CorArray.Length(c2), ActualBytes2, 0, True)
    
    Assert.That ActualLength, Iz.EqualTo(ExpectedLength)
    Assert.That ActualBytes1, Iz.EqualTo(ExpectedBytes1)
    Assert.That ActualBytes2, Iz.EqualTo(ExpectedBytes2)
End Sub

Public Sub GetBytes_WithTrailingHighSurrogateAndFlushing_EncodesFallback()
    Dim Encoder As Encoder
    Dim ActualCount As Long
    Dim ActualBytes(0 To 15) As Byte
    
    Set Encoder = MakeEncoder
    
    ActualCount = Encoder.GetBytes(NewChars("a", "b", &HD800), 0, 3, ActualBytes, 0, True)
    
    Assert.That ActualCount, Iz.EqualTo(16)
    Assert.That ActualBytes, Iz.EqualTo(NewBytes(97, 0, 0, 0, 98, 0, 0, 0, 63, 0, 0, 0, 63, 0, 0, 0))
End Sub

Public Sub GetBytes_WithByteBufferTooSmall_ThrowsArgumentException()
    Dim Encoder As Encoder
    Dim TooSmall(0 To 9) As Byte
        
    On Error Resume Next
    MakeEncoder.GetBytes NewChars("a", "b", &HD800), 0, 3, TooSmall, 0, True
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub Fallback_WhenCalled_ReturnsSameInstanceAsEncoding()
    Dim Encoder As Encoder
    Set Encoder = Encoding.UTF32.GetEncoder
    
    Dim Actual As EncoderFallback
    Set Actual = Encoder.Fallback
    
    Assert.That Actual, Iz.SameAs(Encoding.UTF32.EncoderFallback)
End Sub

Public Sub Fallback_WhenSetToNothing_ThrowsArgumentNullException()
    Dim Encoder As Encoder
    Set Encoder = Encoding.UTF32.GetEncoder
    
    On Error Resume Next
    Set Encoder.Fallback = Nothing
    
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub FallbackBuffer_WhenCalled_ReturnsCachedInstance()
    Dim Encoder As Encoder
    Set Encoder = Encoding.UTF32.GetEncoder
    
    Dim b1 As EncoderFallbackBuffer
    Dim b2 As EncoderFallbackBuffer
    Set b1 = Encoder.FallbackBuffer
    Set b2 = Encoder.FallbackBuffer
    
    Assert.That b1, Iz.Not.Nothing
    Assert.That b1, Iz.SameAs(b2)
End Sub

Public Sub Reset_WhenCalled_ClearsInternalState()
    Dim Encoder As Encoder
    Dim b1()        As Byte
    Dim b2(0 To 3)  As Byte
    Dim Actual      As Long
    
    Set Encoder = Encoding.UTF32.GetEncoder
    
    b1 = NewBytes()
    Actual = Encoder.GetBytes(NewChars(&HD800), 0, 1, b1, 0, False)
    Assert.That Actual, Iz.EqualTo(0)
    
    Encoder.Reset
    
    Actual = Encoder.GetBytes(NewChars(&HDC00), 0, 1, b2, 0, False)
    Assert.That Actual, Iz.EqualTo(4)
    Assert.That b2, Iz.EqualTo(NewBytes(&HFD, &HFF, 0, 0))
End Sub

Public Sub Convert_WithLimitedOutputBuffer_ConvertsWithinLimits()
    Dim Chars(10000)    As Integer
    Dim Bytes(10000)    As Byte
    Dim CharsUsed       As Long
    Dim BytesUsed       As Long
    Dim Completed       As Boolean
    Dim Encoder         As Encoder
    Set Encoder = Encoding.UTF32.GetEncoder
    
    Encoder.Convert Chars, 0, 10000, Bytes, 0, 2000, True, CharsUsed, BytesUsed, Completed
    
    Assert.That CharsUsed, Iz.EqualTo(312)
    Assert.That BytesUsed, Iz.EqualTo(1248)
    Assert.That Completed, Iz.False
End Sub

Private Function MakeEncoder() As Encoder
    Set MakeEncoder = Encoding.GetEncoding("utf-32", Cor.NewEncoderReplacementFallback("??")).GetEncoder
End Function
