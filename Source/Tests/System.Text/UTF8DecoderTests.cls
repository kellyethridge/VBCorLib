VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF8DecoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UTF8DecoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetCharCount_WithValues_ReturnsExpected"
            Test.Use NewBytes(), 0
            Test.Use NewBytes(&H41), 1 ' a
            Test.Use NewBytes(&H41, &H42, &H43), 3 ' abc
            Test.Use NewBytes(&HC0), 1
            Test.Use NewBytes(&HC2, &H80), 1  ' U+0080
            Test.Use NewBytes(&HC0, &H80), 2 ' 2-byte encoding of U+0000 is invalid
            Test.Use NewBytes(223, 191), 1 ' U+07FF
            Test.Use NewBytes(224, 160, 128), 1 ' U+0800
            Test.Use NewBytes(239, 191, 191), 1 ' U+FFFF
            Test.Use NewBytes(224, 159, 191), 2 ' U+07FF
            Test.Use NewBytes(239, 191, 189), 1 ' U+FFFD
            Test.Use NewBytes(&HED, &HA0, &H80), 2 ' U+D800
            Test.Use NewBytes(240, 144, 128, 128), 2 ' U+D800 U+DC00
            Test.Use NewBytes(240, 65), 2
            Test.Use NewBytes(&H80), 1
            Test.Use NewBytes(&HF6, &HB0, &H80, &H80), 4 ' U+D800 U+0000
        
        Case "GetChars_WithSingleCallWithFlush_ReturnsExpected"
            Test.Use NewBytes(), NewIntegers()
            Test.Use NewBytes(&H41), NewChars("A") ' A
            Test.Use NewBytes(&H41, &H42, &H43), NewChars("A", "B", "C") ' ABC
            Test.Use NewBytes(&HC0), NewChars(&HFFFD)
            Test.Use NewBytes(&HC2, &H80), NewChars(&H80)  ' U+0080
            Test.Use NewBytes(&HC0, &H80), NewChars(&HFFFD, &HFFFD) ' 2-byte encoding of U+0000 is invalid
            Test.Use NewBytes(223, 191), NewChars(&H7FF) ' U+07FF
            Test.Use NewBytes(224, 160, 128), NewChars(&H800) ' U+0800
            Test.Use NewBytes(239, 191, 191), NewChars(&HFFFF) ' U+FFFF
            Test.Use NewBytes(224, 159, 191), NewChars(&HFFFD, &HFFFD) ' U+07FF -> U+FFFD U+FFFD
            Test.Use NewBytes(239, 191, 189), NewChars(&HFFFD) ' U+FFFD
            Test.Use NewBytes(&HED, &HA0, &H80), NewChars(&HFFFD, &HFFFD) ' U+D800
            Test.Use NewBytes(240, 144, 128, 128), NewChars(&HD800, &HDC00) ' U+D800 U+DC00
            Test.Use NewBytes(240, 65), NewChars(&HFFFD, &H41)
            Test.Use NewBytes(&H80), NewChars(&HFFFD)
            Test.Use NewBytes(&HF6, &HB0, &H80, &H80), NewChars(&HFFFD, &HFFFD, &HFFFD, &HFFFD) ' U+D800 U+0000
        
        Case "GetChars_WithSingleCallAndNoFlush_ReturnsExpected"
            Test.Use NewBytes(65), NewChars(65)
            Test.Use NewBytes(&HC0), NewChars(&HFFFD)
            Test.Use NewBytes(&HC2), NewChars()
            Test.Use NewBytes(223), NewChars()
            Test.Use NewBytes(224), NewChars()
            Test.Use NewBytes(224, 160), NewChars()
        
        Case "GetChars_WithMultipleCalls_ReturnsExpected"
            Test.Use NewBytes(&HC2), NewBytes(&H80), NewChars(), NewChars(&H80)
            Test.Use NewBytes(224, 160), NewBytes(128), NewChars(), NewChars(&H800)
            Test.Use NewBytes(&H41, &HC2), NewBytes(&H80), NewChars("A"), NewChars(&H80)
            
    End Select
End Sub

Public Sub GetCharCount_WithValues_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Decoder As Decoder
    Set Decoder = Encoding.UTF8.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(Bytes, LBound(Bytes), CorArray.Length(Bytes), True)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithSingleCallWithFlush_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim Decoder As Decoder
    Set Decoder = Encoding.UTF8.GetDecoder
    
    Dim Chars()     As Integer
    Dim CharCount   As Long
    Dim Actual      As Long
    CharCount = Decoder.GetCharCount(Bytes, LBound(Bytes), CorArray.Length(Bytes), True)
    Chars = CorArray.CreateInstance(vbInteger, CharCount)
    Actual = Decoder.GetChars(Bytes, LBound(Bytes), CorArray.Length(Bytes), Chars, 0, True)
    
    Assert.That Actual, Iz.EqualTo(CorArray.Length(Expected))
    Assert.That Chars, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithSingleCallAndNoFlush_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim Decoder As Decoder
    Set Decoder = Encoding.UTF8.GetDecoder
    
    Dim Chars()     As Integer
    Dim CharCount   As Long
    Dim Actual      As Long
    CharCount = Decoder.GetCharCount(Bytes, LBound(Bytes), CorArray.Length(Bytes), False)
    Chars = CorArray.CreateInstance(vbInteger, CharCount)
    Actual = Decoder.GetChars(Bytes, LBound(Bytes), CorArray.Length(Bytes), Chars, 0, False)
    
    Assert.That Actual, Iz.EqualTo(CorArray.Length(Expected))
    Assert.That Chars, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithMultipleCalls_ReturnsExpected(ByRef Bytes1() As Byte, ByRef Bytes2() As Byte, ByRef Expected1() As Integer, ByRef Expected2() As Integer)
    Dim Decoder As Decoder
    Set Decoder = Encoding.UTF8.GetDecoder
    
    Dim Chars1() As Integer
    Dim Chars2() As Integer
    Chars1 = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(Bytes1, 0, CorArray.Length(Bytes1), False))
    Decoder.GetChars Bytes1, 0, CorArray.Length(Bytes1), Chars1, 0, False
    Chars2 = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(Bytes2, 0, CorArray.Length(Bytes2), True))
    Decoder.GetChars Bytes2, 0, CorArray.Length(Bytes2), Chars2, 0, True
    
    Assert.That Chars1, Iz.EqualTo(Expected1)
    Assert.That Chars2, Iz.EqualTo(Expected2)
End Sub








