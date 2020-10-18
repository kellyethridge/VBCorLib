VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF7EncoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UTF7EncoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetBytes_WithSingleCallNoFlush_ReturnsExpected"
            Test.Use NewIntegers(), NewBytes()
            Test.Use NewIntegers(&H41), NewBytes(&H41)
            Test.Use NewIntegers(&H41, &H41), NewBytes(&H41, &H41)
            Test.Use NewIntegers(&H80), NewBytes(&H2B, &H41, &H49)
            Test.Use NewIntegers(&H80, &H80), NewBytes(&H2B, &H41, &H49, &H41, &H41, &H67)
            Test.Use NewIntegers(vbPlusChar), NewBytes(&H2B, &H2D)
            
        Case "GetBytes_WithSingleCallWithFlush_ReturnsExpected"
            Test.Use NewIntegers(), NewBytes()
            Test.Use NewIntegers(&H41), NewBytes(&H41)
            Test.Use NewIntegers(&H41, &H41), NewBytes(&H41, &H41)
            Test.Use NewIntegers(&H80), NewBytes(&H2B, &H41, &H49, &H41, &H2D)
            Test.Use NewIntegers(&H80, &H80), NewBytes(&H2B, &H41, &H49, &H41, &H41, &H67, &H41, &H2D)
            Test.Use NewIntegers(vbPlusChar), NewBytes(&H2B, &H2D)
        
        Case "GetBytes_WithMultipleCalls_ReturnsExpected"
            Test.Use NewChars("A"), NewChars(), NewBytes(&H41), NewBytes(), 1
            Test.Use NewChars("A"), NewChars("B"), NewBytes(&H41), NewBytes(&H42), 2
            Test.Use NewChars(&H80), NewChars(), NewBytes(&H2B, &H41, &H49), NewBytes(&H41, &H2D), 5
            Test.Use NewChars(&H80), NewChars(&H80), NewBytes(&H2B, &H41, &H49), NewBytes(&H41, &H41, &H67, &H41, &H2D), 8
            
    End Select
End Sub

Public Sub GetBytes_WithSingleCallNoFlush_ReturnsExpected(ByRef Chars() As Integer, ByRef ExpectedBytes() As Byte)
    GetBytes_WithSingleCall_ReturnsExpected Chars, False, ExpectedBytes
End Sub

Public Sub GetBytes_WithSingleCallWithFlush_ReturnsExpected(ByRef Chars() As Integer, ByRef ExpectedBytes() As Byte)
    GetBytes_WithSingleCall_ReturnsExpected Chars, True, ExpectedBytes
End Sub

Public Sub GetBytes_WithMultipleCalls_ReturnsExpected(ByRef c1() As Integer, ByRef c2() As Integer, ByRef ExpectedBytes1() As Byte, ByRef ExpectedBytes2() As Byte, ByVal ExpectedLength As Long)
    Dim Encoder As Encoder
    Set Encoder = Encoding.UTF7.GetEncoder
    
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

Private Sub GetBytes_WithSingleCall_ReturnsExpected(ByRef Chars() As Integer, ByVal Flush As Boolean, ByRef ExpectedBytes() As Byte)
    Dim Encoder As Encoder
    Set Encoder = Encoding.UTF7.GetEncoder

    Dim Actual As Long
    Dim Bytes() As Byte
    Bytes = CorArray.CreateInstance(vbByte, Encoder.GetByteCount(Chars, 0, CorArray.Length(Chars), Flush))
    Actual = Encoder.GetBytes(Chars, 0, CorArray.Length(Chars), Bytes, 0, Flush)
    
    Assert.That Actual, Iz.EqualTo(CorArray.Length(ExpectedBytes))
    Assert.That Bytes, Iz.EqualTo(ExpectedBytes)
End Sub

