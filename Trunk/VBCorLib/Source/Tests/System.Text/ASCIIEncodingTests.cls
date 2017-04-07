VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASCIIEncodingTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ASCIIEncodingTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetBytes_WithChars_ReturnsExpected"
            Test.Use vbNullChar, NewBytes(0)
            Test.Use vbTab, NewBytes(9)
            Test.Use Chr$(127), NewBytes(127)
            Test.Use Chr$(128), NewBytes(63)
            Test.Use ChrW$(&HFFFF), NewBytes(63)
            Test.Use "ABC", NewBytes(65, 66, 67)
            Test.Use NewIntegers(96, 97, 98), NewBytes(96, 97, 98)
            Test.Use "", NewBytes()
            Test.Use NewIntegers(), NewBytes()
            
        Case "GetBytes_WithIndexAndCount_ReturnsExpectedSubset"
            Test.Use "abcd", NewBytes(98, 99)
            Test.Use NewIntegers(97, 98, 99, 100), NewBytes(98, 99)
            
        Case "GetByteCount_WithCharacters_ReturnsExpected"
            Test.Use "abc", 3
            Test.Use "", 0
            Test.Use NewIntegers(97, 98, 99, 100), 4
            Test.Use NewIntegers(), 0
            Test.Use NewIntegers(97, 98, 99, 100, 128), 5
            Test.Use t("\xd800\xdc00"), 2
            
        Case "GetByteCount_WhenUsingEmptyFallback_ReturnsExpected"
            Test.Use "abc", 3
            Test.Use "", 0
            Test.Use NewIntegers(97, 98, 99, 100), 4
            Test.Use NewIntegers(), 0
            Test.Use NewIntegers(97, 98, 99, 100, 128), 4
            Test.Use t("\xd800\xdc00"), 0
            
        Case "GetCharCount_WithBytes_ReturnsExpected"
            Test.Use NewBytes(1), 1
            Test.Use NewBytes(1, 2, 3, 4, 5), 5
            Test.Use NewBytes(), 0
            Test.Use NewBytes(1, 2, 3, 4, 160), 5
        
        Case "GetCharCount_WithEmptyFallback_ReturnsExpected"
            Test.Use NewBytes(1), 1
            Test.Use NewBytes(1, 2, 3, 4, 5), 5
            Test.Use NewBytes(), 0
            Test.Use NewBytes(1, 2, 3, 4, 160), 4
        
        Case "GetString_WithBytes_ReturnsExpected"
            Test.Use NewBytes(48, 49, 50), "012"
            Test.Use NewBytes(), ""
            Test.Use NewBytes(255, 65, 128, 0), "?A?" & vbNullChar
            
        Case "GetChars_WithBytes_ReturnsExpected"
            Test.Use NewBytes(48, 49, 50), NewIntegers(48, 49, 50)
            Test.Use NewBytes(), NewIntegers()
            Test.Use NewBytes(255, 65, 128, 0), NewIntegers(63, 65, 63, 0)
        
        Case "GetMaxByteCount_WithEncoderFallbackCharacter_ReturnsExpected"
            Test.Use "?", 101
            Test.Use "??", 202
            Test.Use "", 101
            
        Case "GetMaxCharCount_WithDecoderFallbackCharacter_ReturnsExpected"
            Test.Use "?", 100
            Test.Use "??", 200
            Test.Use "", 100
            
    End Select
End Sub

Public Sub ZeroLengthArrays()
    Dim Enc As New ASCIIEncoding
    Enc.GetCharCount NewBytes()
    Enc.GetChars NewBytes()
    Enc.GetCharCount NewBytes, 0, 0
    Enc.GetChars NewBytes(), 0, 0
    Enc.GetCharsEx NewBytes(), 0, 0, NewIntegers(), 0
    Enc.GetByteCount NewIntegers()
    Enc.GetBytes NewIntegers()
    Enc.GetByteCount NewIntegers(), 0, 0
    Enc.GetBytes NewIntegers(), 0, 0
    Enc.GetBytesEx NewIntegers(), 0, 0, NewBytes(), 0
    Enc.GetByteCount ""
    Enc.GetBytes ""
End Sub

Public Sub GetBytes_WithChars_ReturnsExpected(ByRef Chars As Variant, ByRef Expected() As Byte)
    Dim Encoding As New ASCIIEncoding
    
    Dim Actual() As Byte
    Actual = Encoding.GetBytes(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithIndexAndCount_ReturnsExpectedSubset(ByRef Chars As Variant, ByRef Expected() As Byte)
    Dim Encoding As New ASCIIEncoding
    
    Dim Actual() As Byte
    Actual = Encoding.GetBytes(Chars, 1, 2)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithCharacters_ReturnsExpected(ByRef Chars As Variant, ByVal Expected As Long)
    Dim Encoding As New ASCIIEncoding
    
    Dim Actual As Long
    Actual = Encoding.GetByteCount(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WhenUsingEmptyFallback_ReturnsExpected(ByRef Chars As Variant, ByVal Expected As Long)
    Dim Enc As ASCIIEncoding
    Set Enc = Encoding.ASCII.Clone
    Set Enc.EncoderFallback = Cor.NewEncoderReplacementFallback("")
    
    Dim Actual As Long
    Actual = Enc.GetByteCount(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithIndexAndCount_ReturnsExpected()
    Dim Encoding As New ASCIIEncoding
    
    Dim Actual As Long
    Actual = Encoding.GetByteCount(NewChars("a", "b", "c"), 1, 2)
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub GetBytesEx_WithCharsThatFitIntoOutput_EncodesAllChars()
    Dim Encoding As New ASCIIEncoding
    Dim Bytes() As Byte
    ReDim Bytes(1 To 10)
    
    Dim ActualCount As Long
    ActualCount = Encoding.GetBytesEx("12345", 0, 5, Bytes, 1)
    
    Assert.That ActualCount, Iz.EqualTo(5)
    Assert.That Bytes, Iz.EqualTo(NewBytes(49, 50, 51, 52, 53, 0, 0, 0, 0, 0)).AsCollection
End Sub

Public Sub GetBytesEx_WithSubsetSpecified_EncodesSubset()
    Dim Encoding As New ASCIIEncoding
    Dim Bytes() As Byte
    ReDim Bytes(1 To 10)
    
    Dim ActualCount As Long
    ActualCount = Encoding.GetBytesEx("12345678", 1, 4, Bytes, 1)
    
    Assert.That ActualCount, Iz.EqualTo(4)
    Assert.That Bytes, Iz.EqualTo(NewBytes(50, 51, 52, 53, 0, 0, 0, 0, 0, 0)).AsCollection
End Sub

Public Sub GetBytesEx_WithEmptyCharacters_DoesNotFill()
    Dim Encoding As New ASCIIEncoding
    Dim Bytes() As Byte
    ReDim Bytes(1 To 10)
    
    Dim ActualCount As Long
    ActualCount = Encoding.GetBytesEx("", 0, 0, Bytes, 1)
    
    Assert.That ActualCount, Iz.EqualTo(0)
    Assert.That Bytes, Iz.All.EqualTo(0)
End Sub

Public Sub GetCharCount_WithBytes_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Encoding As New ASCIIEncoding
    
    Dim Actual As Long
    Actual = Encoding.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithEmptyFallback_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Enc As ASCIIEncoding
    Set Enc = Encoding.ASCII.Clone
    Set Enc.DecoderFallback = Cor.NewDecoderReplacementFallback("")
    
    Dim Actual As Long
    Actual = Enc.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetString_WithBytes_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As String)
    Dim Encoding As New ASCIIEncoding
    
    Dim Actual As String
    Actual = Encoding.GetString(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetString_WithSubsetSpecified_ReturnsSubset()
    Dim Encoding As New ASCIIEncoding
    Dim Bytes() As Byte
    Bytes = NewBytes(255, 255, 48, 49, 50, 255, 255)
    
    Dim Actual As String
    Actual = Encoding.GetString(Bytes, 2, 3)
    
    Assert.That Actual, Iz.EqualTo("012")
End Sub

Public Sub GetChars_WithBytes_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim Encoding As New ASCIIEncoding
    
    Dim Actual() As Integer
    Actual = Encoding.GetChars(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithSubsetSpecified_ReturnsSubset()
    Dim Encoding As New ASCIIEncoding
    Dim Bytes() As Byte
    Bytes = NewBytes(255, 255, 48, 49, 50, 255, 255)
    
    Dim Actual() As Integer
    Actual = Encoding.GetChars(Bytes, 2, 3)
    
    Assert.That Actual, Iz.EqualTo(NewIntegers(48, 49, 50))
End Sub

Public Sub GetCharsEx_WithBytes_ReturnsExpected()
    Dim Encoding As New ASCIIEncoding
    Dim Bytes() As Byte
    Bytes = NewBytes(48, 49, 50)
    
    
    Dim ActualChars(0 To 5) As Integer
    Dim ActualCount As Long
    ActualCount = Encoding.GetCharsEx(Bytes, 0, 3, ActualChars, 0)
    
    Assert.That ActualCount, Iz.EqualTo(3)
    Assert.That ActualChars, Iz.EqualTo(NewIntegers(48, 49, 50, 0, 0, 0))
End Sub

Public Sub GetCharsEx_WithSubsetSpecified_ReturnsSubset()
    Dim Encoding As New ASCIIEncoding
    Dim Bytes() As Byte
    Bytes = NewBytes(48, 49, 50, 51, 52)
    
    
    Dim ActualChars(0 To 5) As Integer
    Dim ActualCount As Long
    ActualCount = Encoding.GetCharsEx(Bytes, 2, 2, ActualChars, 1)
    
    Assert.That ActualCount, Iz.EqualTo(2)
    Assert.That ActualChars, Iz.EqualTo(NewIntegers(0, 50, 51, 0, 0, 0))
End Sub

Public Sub GetMaxByteCount_WithEncoderFallbackCharacter_ReturnsExpected(ByVal Replacement As String, ByVal Expected As Long)
    Dim Enc As Encoding
    Set Enc = Encoding.ASCII.Clone
    Set Enc.EncoderFallback = Cor.NewEncoderReplacementFallback(Replacement)
    
    Dim Actual As Long
    Actual = Enc.GetMaxByteCount(100)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetMaxCharCount_WithDecoderFallbackCharacter_ReturnsExpected(ByVal Replacement As String, ByVal Expected As Long)
    Dim Enc As Encoding
    Set Enc = Encoding.ASCII.Clone
    Set Enc.DecoderFallback = Cor.NewDecoderReplacementFallback(Replacement)
    
    Dim Actual As Long
    Actual = Enc.GetMaxCharCount(100)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetMaxByteCount_WithZero_ReturnsOne()
    Dim Actual As Long
    
    Actual = Encoding.ASCII.GetMaxByteCount(0)
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub GetMaxByteCount_WithTooLargeValue_ThrowsArgumentOutOfRangeException()
    Dim Enc As Encoding
    Set Enc = Encoding.ASCII.Clone
    Set Enc.EncoderFallback = Cor.NewEncoderReplacementFallback("??")
    
    On Error Resume Next
    Enc.GetMaxByteCount &H7FFFFFFE
    
    AssertArgumentOutOfRangeException Err, "CharCount"
End Sub

Public Sub GetMaxCharCount_WithTooLargeValue_ThrowsArgumentOutOfRangeException()
    Dim Enc As Encoding
    Set Enc = Encoding.ASCII.Clone
    Set Enc.DecoderFallback = Cor.NewDecoderReplacementFallback("??")
    
    On Error Resume Next
    Enc.GetMaxCharCount &H7FFFFFFE
    
    AssertArgumentOutOfRangeException Err, "ByteCount"
End Sub

Public Sub Serialization_WithDefaultFallbacks_MaintainsReferenceToDefaults()
    Dim Enc As New ASCIIEncoding
    Dim Actual As ASCIIEncoding
    
    Set Actual = Persist(Enc)
    
    Assert.That Actual.EncoderFallback, Iz.SameAs(EncoderFallback.ReplacementFallback)
    Assert.That Actual.DecoderFallback, Iz.SameAs(DecoderFallback.ReplacementFallback)
End Sub








