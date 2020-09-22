VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF8EncoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UTF8EncoderTests
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
            Test.Use NewChars(), 0
            Test.Use NewChars("a"), 1
            Test.Use NewChars(&HD800), 3
            Test.Use NewChars(&HD800, &HDC00), 4
            
        Case "GetBytes_WithSingleCall_ReturnsExpected"
            Test.Use NewChars(), NewBytes()
            Test.Use NewChars("a"), NewBytes(97)
            Test.Use NewChars(&HD800, &HDC00), NewBytes(240, 144, 128, 128)
            
        Case "GetBytes_WithMultipleCalls_ReturnsExpected"
            Test.Use NewChars(), NewChars(), NewBytes(), NewBytes(), 0
            Test.Use NewChars("a"), NewChars(), NewBytes(97), NewBytes(), 1
            Test.Use NewChars(), NewChars("a"), NewBytes(), NewBytes(97), 1
            Test.Use NewChars(&HD800), NewChars(&HDC00), NewBytes(), NewBytes(240, 144, 128, 128), 4
            Test.Use NewChars(&HDC00), NewChars(&HD800, &HDC00), NewBytes(239, 191, 189), NewBytes(240, 144, 128, 128), 7
        
    End Select
End Sub

Public Sub GetByteCount_WithSingleCall_ReturnsExpected(ByRef Chars() As Integer, ByVal Expected As Long)
    Dim Encoder As Encoder
    Dim Actual As Long
    
    Set Encoder = Encoding.UTF8.GetEncoder
    
    Actual = Encoder.GetByteCount(Chars, 0, CorArray.Length(Chars), True)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithTrailingHighSurrogate_DoesNotStoreHighSurrogate()
    Dim Encoder As Encoder
    Dim Actual As Long
    
    Set Encoder = MakeEncoder
    
    Actual = Encoder.GetByteCount(NewChars("a", &HD800), 0, 2, False)
    Assert.That Actual, Iz.EqualTo(1)
    Actual = Encoder.GetByteCount(NewChars(), 0, 0, True)
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub GetByteCount_WithTrailingHighSurrogateAndFlushing_CountsFallbackCharacters()
    Dim Encoder As Encoder
    Dim Actual As Long
    
    Set Encoder = MakeEncoder
    
    Actual = Encoder.GetByteCount(NewChars("a", &HD800), 0, 2, True)
    
    Assert.That Actual, Iz.EqualTo(4)
End Sub

Public Sub GetBytes_WithSingleCall_ReturnsExpected(ByRef Chars() As Integer, ByRef Expected() As Byte)
    Dim Encoder As Encoder
    Dim Actual() As Byte
    
    Set Encoder = MakeEncoder
    Actual = CorArray.CreateInstance(vbByte, CorArray.Length(Expected))
    
    Encoder.GetBytes Chars, 0, CorArray.Length(Chars), Actual, 0, True
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithMultipleCalls_ReturnsExpected(ByRef Chars1() As Integer, ByRef Chars2() As Integer, ByRef Expected1() As Byte, ByRef Expected2() As Byte, ByVal ExpectedSize As Long)
    Dim Encoder As Encoder
    Dim Actual1() As Byte
    Dim Actual2() As Byte
    Dim ActualSize As Long
    
    Set Encoder = MakeEncoder
    
    Actual1 = CorArray.CreateInstance(vbByte, Encoder.GetByteCount(Chars1, 0, CorArray.Length(Chars1), False))
    ActualSize = Encoder.GetBytes(Chars1, 0, CorArray.Length(Chars1), Actual1, 0, False)
    Actual2 = CorArray.CreateInstance(vbByte, Encoder.GetByteCount(Chars2, 0, CorArray.Length(Chars2), True))
    ActualSize = ActualSize + Encoder.GetBytes(Chars2, 0, CorArray.Length(Chars2), Actual2, 0, True)
    
    Assert.That Actual1, Iz.EqualTo(Expected1)
    Assert.That Actual2, Iz.EqualTo(Expected2)
    Assert.That ActualSize, Iz.EqualTo(ExpectedSize)
End Sub

Public Sub GetBytes_WithTrailingHighSurrogateAndFlushing_ThrowsEncoderFallbackException()
    Dim Encoder As Encoder
    Dim ActualBytes(0 To 15) As Byte
    
    Set Encoder = MakeEncoder(True)
    
    On Error Resume Next
    Encoder.GetBytes NewChars("a", "b", &HD800), 0, 3, ActualBytes, 0, True
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(2), "Index"
    Assert.That Ex.CharUnknown, Iz.EqualTo(&HD800), "CharUnknown"
End Sub

Public Sub GetBytes_WithInitialHangingHighSurrogate_ThrowsEncoderFallbackException()
    Dim Encoder As Encoder
    Dim ActualBytes(0 To 15) As Byte
    
    Set Encoder = MakeEncoder(True)
    Encoder.GetBytes NewChars("a", &HD800), 0, 2, ActualBytes, 0, False
    
    On Error Resume Next
    Encoder.GetBytes NewChars("a"), 0, 1, ActualBytes, 0, False
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(-1), "Index"
    Assert.That Ex.CharUnknown, Iz.EqualTo(&HD800), "CharUnknown"
End Sub


Private Function MakeEncoder(Optional ByVal ThrowOnInvalidBytes As Boolean) As Encoder
    Set MakeEncoder = NewUTF8Encoding(False, ThrowOnInvalidBytes).GetEncoder
End Function








