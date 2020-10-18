VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF7DecoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UTF7DecoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetCharCount_WithSingleCallAndFlush_ReturnsExpected"
            Test.Use NewBytes(65, 43, 73, 109, 73, 68, 107, 81, 45, 46), 4  ' "A\x2262\x0391."
            Test.Use NewBytes(72, 105, 32, 77, 111, 109, 32, 45, 43, 74, 106, 111, 45, 45, 43, 65, 67, 69, 45), 11  ' "Hi Mom -\x263A-!"
            Test.Use NewBytes(43, 90, 101, 86, 110, 76, 73, 113, 101, 45), 3    ' "\x65E5\x672C\x8A9E"
            Test.Use NewBytes(72, 105, 32, 77, 111, 109, 32, 43, 74, 106, 111, 65, 73, 81, 45), 9   ' "Hi Mom \x263A!"            Test.Use NewBytes(43), 0
            Test.Use NewBytes(45), 1
            Test.Use NewBytes(43, 128), 1
            Test.Use NewBytes(43, 97, 128, 97, 97), 3
            Test.Use NewBytes(43, 97, 97, 128, 97), 2
            Test.Use NewBytes(43, 97, 97, 128, 97, 43, 97), 2
            Test.Use NewBytes(43, 97, 2), 1
            Test.Use NewBytes(43, 97, 2, 43, 97), 1
            Test.Use NewBytes(43, 97, 45), 0
            Test.Use NewBytes(43, 43, 43), 0
        
        Case "GetCharCount_WithSingleCallAndNoFlush_ReturnsExpected"
            Test.Use NewBytes(65), 1
            Test.Use NewBytes(43), 0
            Test.Use NewBytes(65, 65, 43, 65), 2
            
        Case "GetCharCount_WithMultipleCalls_DoesNotStoreState"
            Test.Use NewBytes(), NewBytes(), 0, 0
            Test.Use NewBytes(65), NewBytes(), 1, 0
            Test.Use NewBytes(), NewBytes(65), 0, 1
            Test.Use NewBytes(65, 43, 73, 109), NewBytes(73, 68, 7, 81, 45, 46), 1, 6
            Test.Use NewBytes(65, 43, 73, 109, 73), NewBytes(68, 7, 43, 65, 46), 2, 3
            
        Case "GetChars_WithSingleCallAndFlush_ReturnsExpected"
            Test.Use NewBytes(45), NewChars(45)
            Test.Use NewBytes(65, 65), NewChars(65, 65)
            Test.Use NewBytes(&H80), NewChars(&H80)
            Test.Use NewBytes(43), NewChars()
            Test.Use NewBytes(43, 45), NewChars(&H2B)
            Test.Use NewBytes(43, 128), NewChars(128)
            Test.Use NewBytes(43, 97, 128, 97, 97), NewChars(128, 97, 97)
            Test.Use NewBytes(43, 97, 97, 128, 97), NewChars(128, 97)
            Test.Use NewBytes(43, 97, 97, 128, 97, 43, 97), NewChars(128, 97)
            Test.Use NewBytes(43, 97, 128, 65, 43, 97, 97), NewChars(128, 65)
            Test.Use NewBytes(43, 43, 43), NewChars()

        Case "GetChars_WithMultipleCalls_ReturnsExpected"
            Test.Use NewBytes(), NewBytes(), NewChars(), NewChars()
            Test.Use NewBytes(65), NewBytes(66), NewChars(65), NewChars(66)
            Test.Use NewBytes(&H41, &H2B, &H49, &H6D), NewBytes(&H49, &H44, &H6B, &H51, &H2D, &H2E), NewChars(&H41), NewChars(&H2262, &H391, &H2E)
            Test.Use NewBytes(&H41, &H2B, &H49, &H6D, &H49), NewBytes(&H44, &H6B, &H51, &H2D, &H2E), NewChars(&H41, &H2262), NewChars(&H391, &H2E)
            Test.Use NewBytes(&H41, &H2B, &H49, &H6D, &H49), NewBytes(&H80, &H41), NewChars(&H41, &H2262), NewChars(&H80, &H41)
            Test.Use NewBytes(&H41, &H2B, &H49, &H6D, &H49), NewBytes(&H80, &H41, &H2B, &H49, &H63, &H49), NewChars(&H41, &H2262), NewChars(&H80, &H41, &H21C2)
        
    End Select
End Sub

Public Sub GetCharCount_WithSingleCallAndFlush_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Decoder As Decoder
    Set Decoder = Encoding.UTF7.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(Bytes, 0, CorArray.Length(Bytes), True)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithSingleCallAndNoFlush_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Decoder As Decoder
    Set Decoder = Encoding.UTF7.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(Bytes, 0, CorArray.Length(Bytes), False)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithMultipleCalls_DoesNotStoreState(ByRef Bytes1() As Byte, ByRef Bytes2() As Byte, ByVal ExpectedCount1 As Long, ByVal ExpectedCount2 As Long)
    Dim Decoder As Decoder
    Set Decoder = Encoding.UTF7.GetDecoder
    
    Dim Actual1 As Long
    Dim Actual2 As Long
    Actual1 = Decoder.GetCharCount(Bytes1, 0, CorArray.Length(Bytes1), False)
    Actual2 = Decoder.GetCharCount(Bytes2, 0, CorArray.Length(Bytes2), True)
    
    Assert.That Actual1, Iz.EqualTo(ExpectedCount1), "#1"
    Assert.That Actual2, Iz.EqualTo(ExpectedCount2), "#2"
End Sub

Public Sub GetChars_WithSingleCallAndFlush_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim Decoder As Decoder
    Set Decoder = Encoding.UTF7.GetDecoder
    
    Dim Chars() As Integer
    Dim Actual  As Long
    Chars = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(Bytes, 0, CorArray.Length(Bytes), True))
    Actual = Decoder.GetChars(Bytes, 0, CorArray.Length(Bytes), Chars, 0, True)
    
    Assert.That Actual, Iz.EqualTo(CorArray.Length(Expected))
    Assert.That Chars, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithMultipleCalls_ReturnsExpected(ByRef Bytes1() As Byte, ByRef Bytes2() As Byte, ByRef ExpectedChars1() As Integer, ByRef ExpectedChars2() As Integer)
    Dim Decoder As Decoder
    Set Decoder = Encoding.UTF7.GetDecoder
    
    Dim Chars1() As Integer
    Dim Chars2() As Integer
    Dim Actual1 As Long
    Dim Actual2 As Long
    Chars1 = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(Bytes1, 0, CorArray.Length(Bytes1), False))
    Actual1 = Decoder.GetChars(Bytes1, 0, CorArray.Length(Bytes1), Chars1, 0, False)
    Chars2 = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(Bytes2, 0, CorArray.Length(Bytes2), True))
    Actual2 = Decoder.GetChars(Bytes2, 0, CorArray.Length(Bytes2), Chars2, 0, True)
    
    Assert.That Chars1, Iz.EqualTo(ExpectedChars1)
    Assert.That Chars2, Iz.EqualTo(ExpectedChars2)
    Assert.That Actual1, Iz.EqualTo(CorArray.Length(Chars1))
    Assert.That Actual2, Iz.EqualTo(CorArray.Length(Chars2))
End Sub

























