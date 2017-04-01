VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnicodeDecoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UnicodeDecoderTests
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
            Test.Use NewBytes(0), 0
            Test.Use NewBytes(0, 0), 1
            Test.Use NewBytes(0, 0, 0), 1
            Test.Use NewBytes(0, 0, 0, 0), 2
            Test.Use NewBytes(&H0, &HD8), 0
            Test.Use NewBytes(&H0, &HD8, &H0), 0
            Test.Use NewBytes(&H0, &HD8, &H0, &HDC), 2
            Test.Use NewBytes(&H0, &HD8, &H0, &H41, &H0, &H42), 3
            
        Case "GetCharCount_WithMultipleCalls_ReturnsExpected"
            Test.Use NewBytes(), NewBytes(), 0
            Test.Use NewBytes(0), NewBytes(), 0
            Test.Use NewBytes(), NewBytes(0), 0
            Test.Use NewBytes(0), NewBytes(0), 1
            Test.Use NewBytes(&H0, &HD8), NewBytes(), 0
            Test.Use NewBytes(&H0, &HD8), NewBytes(&H0), 0
            Test.Use NewBytes(&H0, &HD8), NewBytes(&H0, &HDC), 2
            
    End Select
End Sub

Public Sub GetCharCount_WithValues_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(Bytes, 0, CorArray.Length(Bytes))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithMultipleCalls_ReturnsExpected(ByRef b1() As Byte, ByRef b2() As Byte, ByVal Expected As Long)
    Dim Decoder As Decoder
    Set Decoder = MakeEmptyDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(b1, 0, CorArray.Length(b1))
    Actual = Actual + Decoder.GetCharCount(b2, 0, CorArray.Length(b2))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Private Function MakeEmptyDecoder() As Decoder
    Dim Enc As UnicodeEncoding
    Set Enc = Encoding.Unicode.Clone
    Set Enc.DecoderFallback = Cor.NewDecoderReplacementFallback("")
    
    Set MakeEmptyDecoder = Enc.GetDecoder
End Function

