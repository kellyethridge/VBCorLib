VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBCSCodePageEncodingTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' DBCSCodePageEncodingTests
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
        Case "GetCharCount_WithValues_ReturnsExpected"
            Test.Use NewBytes(), 0
            Test.Use NewBytes(&H41), 1
            Test.Use NewBytes(&H41, &H42), 2
            Test.Use NewBytes(&HFF, &HFF, &HFF), 3
            Test.Use NewBytes(&H81, &H40), 1
            
        Case "GetChars_WithValues_ReturnsExpected"
            Test.Use NewBytes(), NewChars()
            Test.Use NewBytes(&H41), NewChars(&H41)
            Test.Use NewBytes(&H41, &H42), NewChars(&H41, &H42)
            Test.Use NewBytes(&HFF, &HFF, &HFF), NewChars(&HF8F5, &HF8F5, &HF8F5)
            Test.Use NewBytes(&H81, &H40), NewChars(&H4E02)
            Test.Use NewBytes(&H81, &H3F), NewChars("?")
            Test.Use NewBytes(&H81), NewChars("?")
        
        Case "GetByteCount_WithValues_ReturnsExpected"
            Test.Use NewChars(), 0
            Test.Use NewChars(&H41), 1
            Test.Use NewChars(&H41, &H42), 2
            Test.Use NewChars(&HFF, &HFF, &HFF), 3
            Test.Use NewChars(&H4E02), 2
            Test.Use NewChars(&HFFFD), 1
            
        Case "GetBytes_WithValues_ReturnsExpected"
            Test.Use NewChars(), NewBytes()
            Test.Use NewChars(&H41), NewBytes(&H41)
            Test.Use NewChars(&H4E02, &H41), NewBytes(&H81, &H40, &H41)
            
        Case "GetMaxCharCount_WithCustomReplacementFallback_ReturnsExpected"
            Test.Use 0, 3
            Test.Use 1, 6
            Test.Use 10, 33
            
        Case "GetMaxByteCount_WithCustomReplacementFallback_ReturnsExpected"
            Test.Use 0, 6
            Test.Use 1, 12
            Test.Use 10, 66
            
    End Select
End Sub

Public Sub GetCharCount_WithValues_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim En As Encoding
    Set En = MakeEncoding
    
    Dim Actual As Long
    Actual = En.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithInvalidTrailingByte_ThrowsDecoderFallbackException()
    Dim En As Encoding
    Set En = MakeEncoding.Clone
    Set En.DecoderFallback = DecoderFallback.ExceptionFallback
    
    On Error Resume Next
    En.GetCharCount NewBytes(&H81, &H3F)
    
    AssertDecoderFallbackException Err
End Sub

Public Sub GetChars_WithValues_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim En As Encoding
    Set En = MakeEncoding
    
    Dim Actual() As Integer
    Actual = En.GetChars(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharsEx_WithInvalidChar_ThrowsDecoderFallbackException()
    Dim En As Encoding
    Set En = MakeEncoding.Clone
    Set En.DecoderFallback = DecoderFallback.ExceptionFallback
    
    On Error Resume Next
    Dim Actual(0) As Integer
    En.GetCharsEx NewBytes(&H81, &H3F), 0, 2, Actual, 0
    
    AssertDecoderFallbackException Err
End Sub

Public Sub GetCharsEx_WithTooSmallConversionBuffer_ThrowsArgumentException()
    Dim En As Encoding
    Set En = MakeEncoding
    
    On Error Resume Next
    Dim TooSmall(0) As Integer
    En.GetCharsEx NewBytes(&H41, &H42), 0, 2, TooSmall, 0
    
    AssertArgumentException Err, "Chars"
End Sub

Public Sub GetByteCount_WithValues_ReturnsExpected(ByRef Chars() As Integer, ByVal Expected As Long)
    Dim En As Encoding
    Set En = MakeEncoding
    
    Dim Actual As Long
    Actual = En.GetByteCount(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithInvalidChar_ThrowsEncoderFallbackException()
    Dim En As Encoding
    Set En = MakeEncoding.Clone
    Set En.EncoderFallback = EncoderFallback.ExceptionFallback

    On Error Resume Next
    En.GetByteCount NewChars(&HFFFD)
    
    AssertEncoderFallbackException Err
End Sub

Public Sub GetBytes_WithValues_ReturnsExpected(ByRef Chars() As Integer, ByRef Expected() As Byte)
    Dim En As Encoding
    Set En = MakeEncoding
    
    Dim Actual() As Byte
    Actual = En.GetBytes(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithUntranslatableCharacter_ThrowsEncoderFallbackException()
    Dim En As Encoding
    Set En = Encoding.GetEncoding(ChineseSimplified, EncoderFallback.ExceptionFallback)
    
    On Error Resume Next
    Dim Bytes(100) As Byte
    En.GetBytesEx NewChars("A", "b", "C", &H2232), 0, 4, Bytes, 0
    
    AssertEncoderFallbackException Err
End Sub

Public Sub GetMaxCharCount_WithCustomReplacementFallback_ReturnsExpected(ByVal Count As Long, ByVal Expected As Long)
    Dim Enc As Encoding
    Dim Actual As Long
    
    Set Enc = MakeEncoding("abc")
    
    Actual = Enc.GetMaxCharCount(Count)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetMaxByteCount_WithCustomReplacementFallback_ReturnsExpected(ByVal Count As Long, ByVal Expected As Long)
    Dim Enc As Encoding
    Dim Actual As Long
    
    Set Enc = MakeEncoding("abc")
    
    Actual = Enc.GetMaxByteCount(Count)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Private Function MakeEncoding(Optional ByVal Replacement As String) As Encoding
    Set MakeEncoding = Encoding.GetEncoding(ChineseSimplified)
    
    If Len(Replacement) > 0 Then
        Set MakeEncoding = MakeEncoding.Clone
        Set MakeEncoding.DecoderFallback = Cor.NewDecoderReplacementFallback(Replacement)
        Set MakeEncoding.EncoderFallback = Cor.NewEncoderReplacementFallback(Replacement)
    End If
End Function
