VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecoderReplacementFallbackTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' DecoderReplacementFallback tests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource


Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Equals_WithInvalidTypes_ReturnsValues"
            Test.Use Nothing
            Test.Use CObj(Me)
            
    End Select
End Sub

Public Sub DefaultString_WithDefaultConstructor_ReturnsQuestionMark()
    Dim Decoder As New DecoderReplacementFallback
    
    Dim Actual As String
    Actual = Decoder.DefaultString
    
    Assert.That Actual, Iz.EqualTo("?")
End Sub

Public Sub MaxCharCount_WithDefaultConstructor_ReturnsOne()
    Dim Decoder As New DecoderReplacementFallback
    
    Dim Actual As Long
    Actual = Decoder.MaxCharCount
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub DefaultString_WithEmptyString_ReturnsEmptyString()
    Dim Decoder As DecoderReplacementFallback
    Set Decoder = Cor.NewDecoderReplacementFallback("")
    
    Dim Actual As String
    Actual = Decoder.DefaultString
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub MaxCharCount_WithEmptyString_ReturnsZero()
    Dim Decoder As DecoderReplacementFallback
    Set Decoder = Cor.NewDecoderReplacementFallback("")
    
    Dim Actual As Long
    Actual = Decoder.MaxCharCount
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub Equals_WithDifferingDefaultStrings_ReturnsFalse()
    Dim Decoder1 As New DecoderReplacementFallback
    Dim Decoder2 As DecoderReplacementFallback
    Set Decoder2 = Cor.NewDecoderReplacementFallback("!")
    
    Dim Actual As Boolean
    Actual = Decoder1.Equals(Decoder2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equals_WithSameDefaultString_ReturnsTrue()
    Dim Decoder1 As DecoderReplacementFallback
    Dim Decoder2 As DecoderReplacementFallback
    Set Decoder1 = Cor.NewDecoderReplacementFallback("!")
    Set Decoder2 = Cor.NewDecoderReplacementFallback("!")
    
    Dim Actual As Boolean
    Actual = Decoder1.Equals(Decoder2)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Equals_WithInvalidTypes_ReturnsValues(ByRef Value As Variant)
    Dim Decoder As New DecoderReplacementFallback
    
    Dim Actual As Boolean
    Actual = Decoder.Equals(Value)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Constructor_WithHighSurrogatePair_ThrowsArgumentException()
    Dim HighSurrogatePair As String
    HighSurrogatePair = ChrW$(&HD800) & ChrW$(&HDBFF)
    
    On Error Resume Next
    Cor.NewDecoderReplacementFallback HighSurrogatePair
    
    AssertArgumentException Err, "Replacement"
End Sub

Public Sub Constructor_WithLowSurrogatePair_ThrowsArgumentException()
    Dim LowSurrogatePair As String
    LowSurrogatePair = ChrW$(&HDC00) & ChrW$(&HDFFF)
    
    On Error Resume Next
    Cor.NewDecoderReplacementFallback LowSurrogatePair
    
    AssertArgumentException Err, "Replacement"
End Sub

Public Sub Constructor_WithHighLowSurrogatePair_DoesNotThrow()
    Dim SurrogatePair As String
    SurrogatePair = ChrW$(&HD800) & ChrW$(&HDC00)
    
    Dim Actual As DecoderReplacementFallback
    Set Actual = Cor.NewDecoderReplacementFallback(SurrogatePair)
    
    Assert.That Actual.DefaultString, Iz.EqualTo(SurrogatePair)
End Sub

Public Sub Constructor_WithLowHighSurrogatePair_ThrowsArgumentException()
    Dim SurrogatePair As String
    SurrogatePair = ChrW$(&HDC00) & ChrW$(&HD800)
    
    On Error Resume Next
    Cor.NewDecoderReplacementFallback SurrogatePair
    
    AssertArgumentException Err, "Replacement"
End Sub

Public Sub Constructor_WithUnmatchedHighSurrogate_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewDecoderReplacementFallback "a" & ChrW$(&HD800) & "b"
    
    AssertArgumentException Err, "Replacement"
End Sub

Public Sub Constructor_WithUnmatchedLowSurrogate_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewDecoderReplacementFallback "a" & ChrW$(&HDC00) & "b"
    
    AssertArgumentException Err, "Replacement"
End Sub

Public Sub Constructor_WithSeparatedSurrogates_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewDecoderReplacementFallback ChrW$(&HD800) & "a" & ChrW$(&HDC00)
    
    AssertArgumentException Err, "Replacement"
End Sub


