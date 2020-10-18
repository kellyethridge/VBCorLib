VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBCSCodePageDecoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' DBCSCodePageDecoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Const ChineseSimplified As Long = 936

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetCharCount_WithSingleCallAndFlush_ReturnsExpected"
            Test.Use NewBytes(), 0
            Test.Use NewBytes(&H41), 1
            Test.Use NewBytes(&H41, &H42), 2
            Test.Use NewBytes(&H81, &H40), 1
            
        Case "GetChars_WithSingleCallAndFlush_ReturnsExpected"
            Test.Use NewBytes(), NewChars()
            Test.Use NewBytes(&H41), NewChars(&H41)
            Test.Use NewBytes(&H81, &H40), NewChars(&H4E02)
            Test.Use NewBytes(&H81, &H3F), NewChars("?")
            Test.Use NewBytes(&H81), NewChars("?")
            
        Case "GetChars_WithSingleCallAndNoFlush_ReturnsExpected"
            Test.Use NewBytes(), NewChars()
            Test.Use NewBytes(&H41), NewChars(&H41)
            Test.Use NewBytes(&H81, &H40), NewChars(&H4E02)
            Test.Use NewBytes(&H81, &H3F), NewChars("?")
            Test.Use NewBytes(&H81), NewChars()
                    
        Case "GetChars_WithMultipleCalls_ReturnsExpected"
            Test.Use NewBytes(), NewBytes(&H41), NewChars(), NewChars(&H41)
            Test.Use NewBytes(&H81), NewBytes(&H40), NewChars(), NewChars(&H4E02)
            Test.Use NewBytes(&H81), NewBytes(&H40, &H40), NewChars(), NewChars(&H4E02, &H40)
            Test.Use NewBytes(&H81), NewBytes(), NewChars(), NewChars("?")
            
    End Select
End Sub

Public Sub GetCharCount_WithSingleCallAndFlush_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Decoder As Decoder
    Set Decoder = MakeDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(Bytes, 0, CorArray.Length(Bytes), True)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithSingleCallAndFlush_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim Decoder As Decoder
    Set Decoder = MakeDecoder
    
    Dim Actual() As Integer
    Dim ActualLength As Long
    Actual = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(Bytes, 0, CorArray.Length(Bytes), True))
    ActualLength = Decoder.GetChars(Bytes, 0, CorArray.Length(Bytes), Actual, 0, True)
    
    Assert.That ActualLength, Iz.EqualTo(CorArray.Length(Expected))
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithSingleCallAndNoFlush_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim Decoder As Decoder
    Set Decoder = MakeDecoder
    
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
    Set Decoder = MakeDecoder
    
    Dim Chars1() As Integer
    Dim Chars2() As Integer
    Chars1 = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(Bytes1, 0, CorArray.Length(Bytes1), False))
    Decoder.GetChars Bytes1, 0, CorArray.Length(Bytes1), Chars1, 0, False
    Chars2 = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(Bytes2, 0, CorArray.Length(Bytes2), True))
    Decoder.GetChars Bytes2, 0, CorArray.Length(Bytes2), Chars2, 0, True
    
    Assert.That Chars1, Iz.EqualTo(Expected1)
    Assert.That Chars2, Iz.EqualTo(Expected2)
End Sub


Private Function MakeDecoder() As Decoder
    Set MakeDecoder = Encoding.GetEncoding(ChineseSimplified).GetDecoder
End Function
