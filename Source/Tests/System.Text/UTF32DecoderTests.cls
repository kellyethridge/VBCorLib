VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF32DecoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UTF32DecoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetCharCount_WithSingleCallWithoutFlushing_ReturnsExpected"
            Test.Use NewBytes, 0
            Test.Use NewBytes(97, 0, 0, 0), 1
            Test.Use NewBytes(0, 252, 4, 0), 2
            Test.Use NewBytes(0, 0, &H11, 0), 1
            Test.Use NewBytes(0, 0, &H10, 0), 2
            Test.Use NewBytes(&H0, &HD8, 0, 0), 1
            Test.Use NewBytes(&H0, &HDC, 0, 0), 1
            Test.Use NewBytes(97), 0
            Test.Use NewBytes(97, 0), 0
            Test.Use NewBytes(97, 0, 0), 0
        
        Case "GetChars_WithSingleCallAndFlushing_ReturnsExpected"
            Test.Use NewBytes(), NewChars()
            Test.Use NewBytes(97, 0, 0, 0), NewChars("a")
            Test.Use NewBytes(97, 0, 0, 0, 98), NewChars("a", "?", "?")
            
        Case "GetChars_WithSingleCallAndNotFlushing_ReturnsExpected"
            Test.Use NewBytes(), NewChars()
            Test.Use NewBytes(97, 0, 0, 0), NewChars("a")
            Test.Use NewBytes(97, 0, 0, 0, 98), NewChars("a")
        
    End Select
End Sub

Public Sub GetCharCount_WithSingleCallWithoutFlushing_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = Encoding.UTF32.GetDecoder.GetCharCount(Bytes, 0, CorArray.Length(Bytes), False)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithPartialCharsAndFlushing_CountsFallbackCharacters()
    Dim Dec As Decoder
    Dim Actual As Long
    
    Set Dec = MakeDecoder
    
    Actual = Dec.GetCharCount(NewBytes(97, 0, 0, 0, 98, 0, 0), 0, 7, True)
    
    Assert.That Actual, Iz.EqualTo(3)
End Sub

Public Sub GetCharCount_WithPartialCharsAndNotFlushing_IgnoresPartialChars()
    Dim Dec As Decoder
    Dim Actual As Long
    
    Set Dec = MakeDecoder
    
    Actual = Dec.GetCharCount(NewBytes(97, 0, 0, 0, 98, 0, 0, 0, 99, 0, 0), 0, 11, False)
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub GetCharCount_WithMultipleCalls_ReturnsExpected()
    Dim Dec As Decoder
    Dim Bytes1() As Byte
    Dim Bytes2() As Byte
    Dim Chars(0 To 100) As Integer
    
    Set Dec = MakeDecoder
    Bytes1 = NewBytes(97, 0, 0, 0, 98)
    Bytes2 = NewBytes(0, 0, 0)
    
    Assert.That Dec.GetCharCount(Bytes1, 0, 5, False), Iz.EqualTo(1), "#1"
    Dec.GetChars Bytes1, 0, 5, Chars, 0, False
    Assert.That Dec.GetCharCount(Bytes2, 0, 3, True), Iz.EqualTo(1), "#2"
End Sub

Public Sub GetChars_WithSingleCallAndFlushing_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim Dec As Decoder
    Dim ActualLength As Long
    Dim ActualChars() As Integer
    
    ActualChars = CorArray.CreateInstance(vbInteger, CorArray.Length(Expected))
    
    ActualLength = MakeDecoder.GetChars(Bytes, 0, CorArray.Length(Bytes), ActualChars, 0, True)
    
    Assert.That ActualLength, Iz.EqualTo(CorArray.Length(Expected))
    Assert.That ActualChars, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithSingleCallAndNotFlushing_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim Dec As Decoder
    Dim ActualLength As Long
    Dim ActualChars() As Integer
    
    ActualChars = CorArray.CreateInstance(vbInteger, CorArray.Length(Expected))
    
    ActualLength = MakeDecoder.GetChars(Bytes, 0, CorArray.Length(Bytes), ActualChars, 0, False)
    
    Assert.That ActualLength, Iz.EqualTo(CorArray.Length(Expected))
    Assert.That ActualChars, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithMultipleCalls_ReturnsExpected()
    Dim Dec As Decoder
    Dim ActualChars1() As Integer
    Dim ActualChars2() As Integer
    Dim ActualLength1 As Long
    Dim ActualLength2 As Long
    Dim Bytes1() As Byte
    Dim Bytes2() As Byte
    
    Set Dec = MakeDecoder
    Bytes1 = NewBytes(97, 0, 0, 0, 98)
    Bytes2 = NewBytes(0, 0, 0)
    
    ActualChars1 = CorArray.CreateInstance(vbInteger, Dec.GetCharCount(Bytes1, 0, 5))
    ActualLength1 = Dec.GetChars(Bytes1, 0, 5, ActualChars1, 0)
    ActualChars2 = CorArray.CreateInstance(vbInteger, Dec.GetCharCount(Bytes2, 0, 3, True))
    ActualLength2 = Dec.GetChars(Bytes2, 0, 3, ActualChars2, 0, True)
    
    Assert.That ActualLength1, Iz.EqualTo(1), "#1"
    Assert.That ActualChars1, Iz.EqualTo(NewChars("a")), "#2"
    Assert.That ActualLength2, Iz.EqualTo(1), "#3"
    Assert.That ActualChars2, Iz.EqualTo(NewChars("b")), "#4"
End Sub

Private Function MakeDecoder() As Decoder
    Set MakeDecoder = Encoding.GetEncoding("utf-32", , Cor.NewDecoderReplacementFallback("??")).GetDecoder
End Function
