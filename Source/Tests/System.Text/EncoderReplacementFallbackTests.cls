VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderReplacementFallbackTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EncoderReplacementFallbackTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Equals_WithInvalidValues_ReturnsFalse"
            Test.Use Nothing
            Test.Use CObj(Me)
            Test.Use 1
            
    End Select
End Sub

Public Sub DefaultString_WithDefaultConstructor_ReturnsQuestionMark()
    Dim Encoder As New EncoderReplacementFallback
    
    Dim Actual As String
    Actual = Encoder.DefaultString
    
    Assert.That Actual, Iz.EqualTo("?")
End Sub

Public Sub MaxCharCount_WithDefaultConstructor_ReturnsOne()
    Dim Encoder As New EncoderReplacementFallback
    
    Dim Actual As Long
    Actual = Encoder.MaxCharCount
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub DefaultString_WithEmptyString_ReturnsEmptyString()
    Dim Encoder As EncoderReplacementFallback
    Set Encoder = Cor.NewEncoderReplacementFallback("")
    
    Dim Actual As String
    Actual = Encoder.DefaultString
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub MaxCharCount_WithEmptyString_ReturnsZero()
    Dim Encoder As EncoderFallback
    Set Encoder = Cor.NewEncoderReplacementFallback("")
    
    Dim Actual As Long
    Actual = Encoder.MaxCharCount
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub Equals_WithDifferentDefaultStrings_ReturnsFalse()
    Dim Encoder1 As New EncoderReplacementFallback
    Dim Encoder2 As EncoderFallback
    Set Encoder2 = Cor.NewEncoderReplacementFallback("!")
    
    Dim Actual As Boolean
    Actual = Encoder1.Equals(Encoder2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equals_WithEqualDefaultStrings_ReturnsTrue()
    Dim Encoder1 As New EncoderReplacementFallback
    Dim Encoder2 As EncoderFallback
    Set Encoder2 = Cor.NewEncoderReplacementFallback("?")
    
    Dim Actual As Boolean
    Actual = Encoder1.Equals(Encoder2)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Equals_WithInvalidValues_ReturnsFalse(ByRef InvalidValue As Variant)
    Dim Encoder As New EncoderReplacementFallback
    
    Dim Actual As Boolean
    Actual = Encoder.Equals(InvalidValue)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Constructor_WithHighSurrogatePair_ThrowsArgumentException()
    Dim HighSurrogatePair As String
    HighSurrogatePair = ChrW$(&HD800) & ChrW$(&HDBFF)
    
    On Error Resume Next
    Cor.NewEncoderReplacementFallback HighSurrogatePair
    
    AssertArgumentException Err, "Replacement"
End Sub

Public Sub Constructor_WithLowSurrogatePair_ThrowsArgumentException()
    Dim LowSurrogatePair As String
    LowSurrogatePair = ChrW$(&HDC00) & ChrW$(&HDFFF)
    
    On Error Resume Next
    Cor.NewEncoderReplacementFallback LowSurrogatePair
    
    AssertArgumentException Err, "Replacement"
End Sub

Public Sub Constructor_WithHighLowSurrogatePair_DoesNotThrow()
    Dim SurrogatePair As String
    SurrogatePair = ChrW$(&HD800) & ChrW$(&HDC00)
    
    Dim Actual As EncoderReplacementFallback
    Set Actual = Cor.NewEncoderReplacementFallback(SurrogatePair)
    
    Assert.That Actual.DefaultString, Iz.EqualTo(SurrogatePair)
End Sub

Public Sub Constructor_WithLowHighSurrogatePair_ThrowsArgumentException()
    Dim SurrogatePair As String
    SurrogatePair = ChrW$(&HDC00) & ChrW$(&HD800)
    
    On Error Resume Next
    Cor.NewEncoderReplacementFallback SurrogatePair
    
    AssertArgumentException Err, "Replacement"
End Sub

Public Sub Constructor_WithUnmatchedHighSurrogate_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewEncoderReplacementFallback "a" & ChrW$(&HD800) & "b"
    
    AssertArgumentException Err, "Replacement"
End Sub

Public Sub Constructor_WithUnmatchedLowSurrogate_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewEncoderReplacementFallback "a" & ChrW$(&HDC00) & "b"
    
    AssertArgumentException Err, "Replacement"
End Sub

Public Sub Constructor_WithSeparatedSurrogates_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewEncoderReplacementFallback ChrW$(&HD800) & "a" & ChrW$(&HDC00)
    
    AssertArgumentException Err, "Replacement"
End Sub














