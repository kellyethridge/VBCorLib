VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF32EncodingTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UTF32EncodingTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Const MaxSurrogatePair As Long = &H10FFFF

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetPreamble_WhenConstructedWithValues_ReturnsExpected"
            Test.Use False, False, Cor.NewBytes()
            Test.Use False, True, Cor.NewBytes(&HFF, &HFE, &H0, &H0)
            Test.Use True, False, Cor.NewBytes()
            Test.Use True, True, Cor.NewBytes(&H0, &H0, &HFE, &HFF)
            
        Case "GetByteCount_WithChars_ReturnsExpected"
            Test.Use NewChars("a", &HD8FF, &H1FF), 16
            Test.Use NewChars("a", &HD8FF, &HD8FF, &HDC00, "b"), 20
            Test.Use NewChars("a", &HDC00, &H1FF), 16
            Test.Use NewChars("a", &HD8FF), 12
            Test.Use "", 0
        
        Case "GetByteCount_withInvalidSurrogates_CountsFallbackCharacters"
            Test.Use t("\uD8FF"), 4
            Test.Use t("\uD8FF\uD8FF"), 8
            Test.Use t("\uD8FF\uDC00\uD8FF"), 8
            Test.Use t("\uDC00\uD8FF\uDC00\uD8FF"), 12
            
        Case "GetByteCount_withInvalidSurrogatesAndSurrogatePairFallbackCharacters_CountsFallbackCharacters"
            Test.Use t("\uD8FF"), 4
            Test.Use t("\uD8FF\uD8FF"), 8
            Test.Use t("\uD8FF\uDC00\uD8FF"), 8
            Test.Use t("\uDC00\uD8FF\uDC00\uD8FF"), 12
            
        Case "GetByteCount_WhenUsingExceptionFallbackWithValues_ThrowsEncoderFallbackException"
            Test.Use "\u00C4 \uD802\u0033 \u00AE", 2, &HD802
            Test.Use "\u00C4 \uDC02\u0033 \u00AE", 2, &HDC02
            Test.Use " \uD802", 1, &HD802
            Test.Use " \uDC02", 1, &HDC02
            Test.Use "a\uDC02a", 1, &HDC02
            Test.Use "a\uD802a", 1, &HD802
                    
        Case "GetBytes_AsLittleEndianWithChars_ReturnsEncodedChars"
            Test.Use "a", NewBytes(97, 0, 0, 0)
            Test.Use NewChars("a", &HD8FF, &HDC00), NewBytes(97, 0, 0, 0, 0, 252, 4, 0)
            Test.Use NewChars("a", &HD8FF, "b"), NewBytes(97, 0, 0, 0, &HFD, &HFF, 0, 0, 98, 0, 0, 0)
            Test.Use NewChars(&H8000), NewBytes(&H0, &H80, &H0, &H0)
            Test.Use NewChars(&HD8FF, "a", &HDC00), NewBytes(&HFD, &HFF, 0, 0, 97, 0, 0, 0, &HFD, &HFF, 0, 0)
            
        Case "GetBytes_AsBigEndianWithChars_ReturnsEncodedChars"
            Test.Use "a", NewBytes(0, 0, 0, 97)
            Test.Use NewChars("a", &HD8FF, &HDC00), NewBytes(0, 0, 0, 97, 0, 4, 252, 0)
            Test.Use NewChars("a", &HD8FF, "b"), NewBytes(0, 0, 0, 97, 0, 0, &HFF, &HFD, 0, 0, 0, 98)
            Test.Use NewChars(&H8000), NewBytes(&H0, &H0, &H80, &H0)
                    
        Case "GetBytes_WithInvalidSurrogates_IncludesFallbackCharacters"
            Test.Use t("\uD8FF"), NewBytes(&HFD, &HFF, 0, 0)
            Test.Use t("\uD8FF\uD8FF"), NewBytes(&HFD, &HFF, 0, 0, &HFD, &HFF, 0, 0)
            Test.Use t("\uD8FF\uDC00\uD8FF"), NewBytes(&H0, &HFC, &H4, &H0, &HFD, &HFF, &H0, &H0)
            Test.Use t("\uDC00\uD8FF\uDC00\uD8FF"), NewBytes(&HFD, &HFF, &H0, &H0, &H0, &HFC, &H4, &H0, &HFD, &HFF, &H0, &H0)
            
        Case "GetMaxCharCount_WithByteCount_ReturnsMaxCharCount"
            Test.Use 0, 2
            Test.Use 1, 2
            Test.Use 2, 3
            Test.Use 3, 3
            Test.Use 4, 4
        
        Case "GetMaxCharCount_WhenUsingCustomDecoderFallback_IncludesFallbackCharacters"
            Test.Use 0, 3
            Test.Use 1, 3
            Test.Use 2, 4
            Test.Use 3, 4
            Test.Use 4, 6
            
        Case "GetCharCount_WithBytes_ReturnsExpectedCount"
            Test.Use NewBytes, 0
            Test.Use NewBytes(97, 0, 0, 0), 1
            Test.Use NewBytes(0, 252, 4, 0), 2
            Test.Use NewBytes(0, 0, &H11, 0), 1
            Test.Use NewBytes(0, 0, &H10, 0), 2
            Test.Use NewBytes(&H0, &HD8, 0, 0), 1
            Test.Use NewBytes(&H0, &HDC, 0, 0), 1
            Test.Use NewBytes(97), 1
            Test.Use NewBytes(97, 0), 1
            Test.Use NewBytes(97, 0, 0), 1
            
        Case "GetCharCount_WithCustomDecoderFallback_CountsFallbackCharacters"
            Test.Use NewBytes, 0
            Test.Use NewBytes(97, 0, 0, 0), 1
            Test.Use NewBytes(0, 252, 4, 0), 2
            Test.Use NewBytes(0, 0, &H11, 0), 3
            Test.Use NewBytes(0, 0, &H10, 0), 2
            Test.Use NewBytes(&H0, &HD8, 0, 0), 3
            Test.Use NewBytes(&H0, &HDC, 0, 0), 3
            Test.Use NewBytes(97), 3
            Test.Use NewBytes(97, 0), 3
            Test.Use NewBytes(97, 0, 0), 3
        
        Case "GetCharCount_WithSurrogateDecoderFallback_CountsFallbackCharacters"
            Test.Use NewBytes, 0
            Test.Use NewBytes(97, 0, 0, 0), 1
            Test.Use NewBytes(0, 252, 4, 0), 2
            Test.Use NewBytes(0, 0, &H11, 0), 2
            Test.Use NewBytes(0, 0, &H10, 0), 2
            Test.Use NewBytes(&H0, &HD8, 0, 0), 2
            Test.Use NewBytes(&H0, &HDC, 0, 0), 2
            Test.Use NewBytes(97), 2
            Test.Use NewBytes(97, 0), 2
            Test.Use NewBytes(97, 0, 0), 2
            
        Case "GetCharCount_WithBigEndianBytes_ReturnsExpected"
            Test.Use NewBytes, 0
            Test.Use NewBytes(0, 0, 0, 97), 1
            Test.Use NewBytes(0, 0, 0, 97, 0, 0, 0, 98), 2
            Test.Use NewBytes(0, 4, 252, 0), 2
            Test.Use NewBytes(0, &H11, 0, 0), 3
            Test.Use NewBytes(0, &H10, 0, 0), 2
            Test.Use NewBytes(0, 0, &HD8, 0), 3
            Test.Use NewBytes(0, 0, &HDC, 0), 3
            Test.Use NewBytes(97), 3
            Test.Use NewBytes(97, 0), 3
            Test.Use NewBytes(97, 0, 0), 3
            
        Case "GetCharCount_WithInvalidByteSequence_ThrowsDecoderFallbackException"
            Test.Use NewBytes(&HFF, &HFF, &H11, &H0), 0, NewBytes(&HFF, &HFF, &H11, &H0)
            Test.Use NewBytes(&H1, &H0, &H0, &H0, &HFF, &HFF, &H11, &H0), 4, NewBytes(&HFF, &HFF, &H11, &H0)
            Test.Use NewBytes(&HFF, &HFF, &H1), 0, NewBytes(&HFF, &HFF, &H1)
            Test.Use NewBytes(&H1, &H0, &H0, &H0, &HFF, &HFF, &H1), 4, NewBytes(&HFF, &HFF, &H1)
            
        Case "GetCharCount_WithInvalidByteSequenceAndBigEndian_ThrowsDecoderFallbackException"
            Test.Use NewBytes(&H0, &H0, &H0, &H1, &H0, &H11, &HFF, &HFF), 4, NewBytes(&H0, &H11, &HFF, &HFF)
            Test.Use NewBytes(&H0, &H0, &H0, &H1, &H1, &HFF, &HFF), 4, NewBytes(&H1, &HFF, &HFF)
            
        Case "GetCharsEx_WithInvalidByteSequence_ThrowsDecoderFallbackException"
            Test.Use NewBytes(&HFF, &HFF, &H11, &H0), 0, NewBytes(&HFF, &HFF, &H11, &H0)
            Test.Use NewBytes(&H1, &H0, &H0, &H0, &HFF, &HFF, &H11, &H0), 4, NewBytes(&HFF, &HFF, &H11, &H0)
            Test.Use NewBytes(&HFF, &HFF, &H1), 0, NewBytes(&HFF, &HFF, &H1)
            Test.Use NewBytes(&H1, &H0, &H0, &H0, &HFF, &HFF, &H1), 4, NewBytes(&HFF, &HFF, &H1)
            
        Case "GetCharsEx_WithInvalidByteSequenceAndBigEndian_ThrowsDecoderFallbackException"
            Test.Use NewBytes(&H0, &H0, &H0, &H1, &H0, &H11, &HFF, &HFF), 4, NewBytes(&H0, &H11, &HFF, &HFF)
            Test.Use NewBytes(&H0, &H0, &H0, &H1, &H1, &HFF, &HFF), 4, NewBytes(&H1, &HFF, &HFF)
            
        Case "GetCharsEx_WithBytes_ReturnsExpected"
            Test.Use NewBytes, NewChars
            Test.Use NewBytes(97, 0, 0, 0), NewChars("a")
            Test.Use NewBytes(97, 0, 0, 0, 98, 0, 0, 0), NewChars("a", "b")
            Test.Use NewBytes(0, 0, &H11, 0), NewChars(&HFFFD)
            Test.Use NewBytes(0, 0, &HD8, 0), NewChars(&HFFFD)
            Test.Use NewBytes(0, 0, &HDC, 0), NewChars(&HFFFD)
            Test.Use NewBytes(0, 0, 1, 0), NewChars(&HD800, &HDC00)
            Test.Use NewBytes(&HFF, &HFF, &H10, 0), NewChars(&HDBFF, &HDFFF)
            Test.Use NewBytes(97, 0, 0, 0, 97), NewChars("a", &HFFFD)
            Test.Use NewBytes(97, 0, 0, 0, 97, 0), NewChars("a", &HFFFD)
            Test.Use NewBytes(97, 0, 0, 0, 97, 0, 0), NewChars("a", &HFFFD)
            
        Case "GetBytesEx_WithInvalidSurrogate_ThrowsEncoderFallbackException"
            Test.Use "\u00C4 \uD802\u0033 \u00AE", 2, &HD802
            Test.Use "\u00C4 \uDC02\u0033 \u00AE", 2, &HDC02
            Test.Use " \uD802", 1, &HD802
            Test.Use " \uDC02", 1, &HDC02
            
    End Select
End Sub

Public Sub Properties_WhenCalled_ReturnsExpected()
    Dim Enc As New UTF32Encoding
    
    Assert.That Enc.CodePage, Iz.EqualTo(12000), "#1"
    Assert.That Enc.IsBrowserDisplay, Iz.False, "#2"
    Assert.That Enc.IsBrowserSave, Iz.False, "#3"
    Assert.That Enc.IsMailNewsDisplay, Iz.False, "#4"
    Assert.That Enc.IsMailNewsSave, Iz.False, "#5"
    Assert.That Enc.IsSingleByte, Iz.False, "#6"
    Assert.That Enc.IsReadOnly, Iz.True, "#7"
End Sub

Public Sub CodePage_WhenConstructedAsBigEndian_Returns12001()
    Dim Enc As UTF32Encoding
    Dim Actual As Long
    
    Set Enc = Cor.NewUTF32Encoding(True, False)
    
    Actual = Enc.CodePage
    
    Assert.That Actual, Iz.EqualTo(12001)
End Sub

Public Sub GetPreamble_WhenConstructedWithValues_ReturnsExpected(ByVal BigEndian As Boolean, ByVal ByteOrderMark As Boolean, ByRef Expected() As Byte)
    Dim Enc As UTF32Encoding
    Dim Actual() As Byte
    
    Set Enc = Cor.NewUTF32Encoding(BigEndian, ByteOrderMark)
    
    Actual = Enc.GetPreamble
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WhenCalledAsLittleEndian_ReturnsExpected()
    Dim Chars() As Integer
    Dim Enc As UTF32Encoding
    
    Chars = Cor.NewChars("z", "a", &H306, &H1FD, &H3B2, &HD8FF, &HDCFF)
    Set Enc = Cor.NewUTF32Encoding(False, True)
    
    Assert.That Enc.GetByteCount(Chars), Iz.EqualTo(24), "#1"
    Assert.That Enc.GetByteCount(Cor.NewChars), Iz.EqualTo(0), "#2"
End Sub

Public Sub GetByteCount_WhenCalledAsBigEndian_ReturnsExpected()
    Dim Chars() As Integer
    Dim Enc As UTF32Encoding
    
    Chars = Cor.NewChars("z", "a", &H306, &H1FD, &H3B2, &HD8FF, &HDCFF)
    Set Enc = Cor.NewUTF32Encoding(True, True)
    
    Assert.That Enc.GetByteCount(Chars), Iz.EqualTo(24), "#1"
    Assert.That Enc.GetByteCount(Cor.NewChars), Iz.EqualTo(0), "#2"
End Sub

Public Sub GetByteCount_WithChars_ReturnsExpected(ByRef Chars As Variant, ByVal Expected As Long)
    Dim Enc As UTF32Encoding
    Dim Actual As Long
    
    Set Enc = Cor.NewUTF32Encoding(False, True).Clone
    Set Enc.EncoderFallback = Cor.NewEncoderReplacementFallback("ab") ' 8 bytes
    
    Actual = Enc.GetByteCount(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_withInvalidSurrogates_CountsFallbackCharacters(ByRef InvalidSequence As Variant, ByVal Expected As Long)
    Dim Enc As New UTF32Encoding
    Dim Actual As Long
    
    Actual = Enc.GetByteCount(InvalidSequence)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_withInvalidSurrogatesAndSurrogatePairFallbackCharacters_CountsFallbackCharacters(ByRef InvalidSurrogates As Variant, ByVal Expected As Long)
    Dim Enc As UTF32Encoding
    Dim Actual As Long
    
    Set Enc = Cor.NewUTF32Encoding(False, False).Clone
    Set Enc.EncoderFallback = Cor.NewEncoderReplacementFallback(t("\uD800\uDC00"))
    
    Actual = Enc.GetByteCount(InvalidSurrogates)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WhenUsingExceptionFallbackWithValues_ThrowsEncoderFallbackException(ByRef Chars As String, ByVal ExpectedIndex As Long, ByVal ExpectedChar As Integer)
    Dim Enc As Encoding
    Set Enc = NewUTF32Encoding(False, False, True)
    
    On Error Resume Next
    Enc.GetByteCount t(Chars)
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.CharUnknown, Iz.EqualTo(ExpectedChar), "CharUnknown"
End Sub

Public Sub GetByteCount_WithIndexAndCount_ReturnsExpected()
    Dim Enc As New UTF32Encoding
    Dim Actual As Long
    
    Actual = Enc.GetByteCount(t("1\uD80034"), 1, 2)
    
    Assert.That Actual, Iz.EqualTo(8)
End Sub

Public Sub GetBytes_AsLittleEndianWithChars_ReturnsEncodedChars(ByRef Chars As Variant, ByRef Expected() As Byte)
    Dim Enc As New UTF32Encoding
    Dim Actual() As Byte
    
    Actual = Enc.GetBytes(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_AsBigEndianWithChars_ReturnsEncodedChars(ByRef Chars As Variant, ByRef Expected() As Byte)
    Dim Enc As UTF32Encoding
    Dim Actual() As Byte
    
    Set Enc = Cor.NewUTF32Encoding(True, False)
    
    Actual = Enc.GetBytes(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithInvalidSurrogates_IncludesFallbackCharacters(ByRef InvalidSurrogates As Variant, ByRef Expected() As Byte)
    Dim Enc As New UTF32Encoding
    Dim Actual() As Byte
    
    Actual = Enc.GetBytes(InvalidSurrogates)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytesEx_WithTrailingHighSurrogate_EncodesFallback()
    Dim Enc As UTF32Encoding
    Dim ActualBytes(0 To 15) As Byte
    Dim ActualCount As Long
    
    Set Enc = MakeEncoding
    
    ActualCount = Enc.GetBytesEx(t("ab\uD800"), 0, 3, ActualBytes, 0)
    
    Assert.That ActualCount, Iz.EqualTo(16)
    Assert.That ActualBytes, Iz.EqualTo(NewBytes(97, 0, 0, 0, 98, 0, 0, 0, 63, 0, 0, 0, 63, 0, 0, 0))
End Sub

Public Sub GetBytesEx_WithHangingHighSurrogate_EncodesFallback()
    Dim Enc As UTF32Encoding
    Dim ActualBytes(0 To 15) As Byte
    Dim ActualCount As Long
    
    Set Enc = MakeEncoding
    
    ActualCount = Enc.GetBytesEx(t("a\uD800b"), 0, 3, ActualBytes, 0)
    
    Assert.That ActualCount, Iz.EqualTo(16)
    Assert.That ActualBytes, Iz.EqualTo(NewBytes(97, 0, 0, 0, 63, 0, 0, 0, 63, 0, 0, 0, 98, 0, 0, 0))
End Sub

Public Sub GetBytesEx_WhenCalled_ReturnsEncodedBytes()
    Dim Enc As New UTF32Encoding
    Dim ActualBytes(0 To 10) As Byte
    Dim ActualCount As Long
    
    ActualCount = Enc.GetBytesEx("abcd", 1, 2, ActualBytes, 1)
    
    Assert.That ActualCount, Iz.EqualTo(8)
    Assert.That ActualBytes, Iz.EqualTo(NewBytes(0, 98, 0, 0, 0, 99, 0, 0, 0, 0, 0))
End Sub

Public Sub GetBytesEx_WithInvalidSurrogate_ThrowsEncoderFallbackException(ByVal Invalid As String, ByVal ExpectedIndex As Long, ByVal ExpectedChar As Integer)
    Dim Enc As Encoding
    Dim Bytes() As Byte
    
    Set Enc = NewUTF32Encoding(False, False, True)
    Invalid = t(Invalid)
    ReDim Bytes(0 To Enc.GetMaxByteCount(Len(Invalid)))
    
    On Error Resume Next
    Enc.GetBytesEx Invalid, 0, Len(Invalid), Bytes, 0
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.CharUnknown, Iz.EqualTo(ExpectedChar), "CharUnknown"
End Sub

Public Sub GetMaxByteCount_WhenCalled_ReturnsCount()
    Dim Enc As New UTF32Encoding
    Dim Actual As Long
    
    Actual = Enc.GetMaxByteCount(100)
    
    Assert.That Actual, Iz.EqualTo(404)
End Sub

Public Sub GetMaxByteCount_WhenUsingCustomEncoderFallback_IncludesFallbackCharacters()
    Dim Enc As UTF32Encoding
    Dim Actual As Long
    
    Set Enc = Cor.NewUTF32Encoding(False, True).Clone
    Set Enc.EncoderFallback = Cor.NewEncoderReplacementFallback("ab")
    
    Actual = Enc.GetMaxByteCount(100)
    
    Assert.That Actual, Iz.EqualTo(808)
End Sub

Public Sub GetMaxCharCount_WithByteCount_ReturnsMaxCharCount(ByVal ByteCount As Long, ByVal Expected As Long)
    Dim Enc As New UTF32Encoding
    Dim Actual As Long

    Actual = Enc.GetMaxCharCount(ByteCount)

    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetMaxCharCount_WhenUsingCustomDecoderFallback_IncludesFallbackCharacters(ByVal ByteCount As Long, ByVal Expected As Long)
    Dim Enc As UTF32Encoding
    Dim Actual As Long
    
    Set Enc = Encoding.GetEncoding("utf-32", , Cor.NewDecoderReplacementFallback("???"))
    
    Actual = Enc.GetMaxCharCount(ByteCount)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithBytes_ReturnsExpectedCount(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = Encoding.UTF32.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithCustomDecoderFallback_CountsFallbackCharacters(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Actual As Long
    Dim Enc As UTF32Encoding
    
    Set Enc = Encoding.GetEncoding("utf-32", , Cor.NewDecoderReplacementFallback("???"))
    
    Actual = Enc.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithSurrogateDecoderFallback_CountsFallbackCharacters(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Actual As Long
    Dim Enc As UTF32Encoding
    
    Set Enc = Encoding.GetEncoding("utf-32", , Cor.NewDecoderReplacementFallback(t("\uD800\uDC00")))
    
    Actual = Enc.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithBigEndianBytes_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Actual As Long
    Dim Enc As UTF32Encoding
    
    Set Enc = MakeEncoding("utf-32be")
    
    Actual = Enc.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithInvalidByteSequence_ThrowsDecoderFallbackException(ByRef Sequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Enc As UTF32Encoding
    Set Enc = NewUTF32Encoding(False, False, True)
    
    On Error Resume Next
    Enc.GetCharCount Sequence
    
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetCharCount_WithInvalidByteSequenceAndBigEndian_ThrowsDecoderFallbackException(ByRef Sequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Enc As UTF32Encoding
    Set Enc = NewUTF32Encoding(True, False, True)
    
    On Error Resume Next
    Enc.GetCharCount Sequence
    
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetCharsEx_WithBytes_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim ActualLength As Long
    Dim ActualChars() As Integer
    
    ActualChars = CorArray.CreateInstance(vbInteger, CorArray.Length(Expected))
    
    ActualLength = Encoding.UTF32.GetCharsEx(Bytes, LBound(Bytes), CorArray.Length(Bytes), ActualChars, LBound(Expected))
    
    Assert.That ActualLength, Iz.EqualTo(CorArray.Length(Expected)), "Length"
    Assert.That ActualChars, Iz.EqualTo(Expected), "Chars"
End Sub

Public Sub GetCharsEx_WithInvalidByteSequence_ThrowsDecoderFallbackException(ByRef Sequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Enc As UTF32Encoding
    Dim Chars(0 To 100) As Integer
    
    Set Enc = NewUTF32Encoding(False, False, True)
    
    On Error Resume Next
    Enc.GetCharsEx Sequence, LBound(Sequence), CorArray.Length(Sequence), Chars, 0
    
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetCharsEx_WithInvalidByteSequenceAndBigEndian_ThrowsDecoderFallbackException(ByRef Sequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Enc As UTF32Encoding
    Dim Chars(0 To 100) As Integer
    
    Set Enc = NewUTF32Encoding(True, False, True)
    
    On Error Resume Next
    Enc.GetCharsEx Sequence, LBound(Sequence), CorArray.Length(Sequence), Chars, 0
    
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetChars_WhenCalled_ReturnsExpected()
    Dim Actual() As Integer
    
    Actual = Encoding.UTF32.GetChars(NewBytes(97, 0, 0, 0))
    
    Assert.That Actual, Iz.EqualTo(NewChars("a"))
End Sub

Public Sub GetString_WhenCalled_ReturnsExpected()
    Dim Actual As String
    
    Actual = Encoding.UTF32.GetString(NewBytes(97, 0, 0, 0))
    
    Assert.That Actual, Iz.EqualTo("a")
End Sub

Private Function MakeEncoding(Optional ByVal EncodingName As String = "utf-32") As UTF32Encoding
    Set MakeEncoding = Encoding.GetEncoding(EncodingName, Cor.NewEncoderReplacementFallback("??"), Cor.NewDecoderReplacementFallback("???"))
End Function
