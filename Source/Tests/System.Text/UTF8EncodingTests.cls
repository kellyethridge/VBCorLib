VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF8EncodingTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UTF8EncodingTests
'
' Tests are replicated from the Mono Project: https://github.com/mono/mono/blob/master/mcs/class/corlib/Test/System.Text/UTF8EncodingTest.cs
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private mUTF8 As UTF8Encoding

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "T2_Boundary_1_FirstPossibleSequence_Pass"
            Test.Use NewBytes(&H0), &H0
            Test.Use NewBytes(&HC2, &H80), &H80
            Test.Use NewBytes(&HE0, &HA0, &H80), &H800
            Test.Use NewBytes(&HF0, &H90, &H80, &H80), &HDC00D800
            
        Case "T2_Boundary_2_LastPossibleSequence_Pass"
            Test.Use NewBytes(&H7F), &H7F&
            Test.Use NewBytes(&HDF, &HBF), &H7FF&
            Test.Use NewBytes(&HEF, &HBF, &HBF), &HFFFF&

        Case "T2_Boundary_3_Other_Pass"
            Test.Use NewBytes(&HED, &H9F, &HBF), &HD7FF&
            Test.Use NewBytes(&HEE, &H80, &H80), &HE000&
            Test.Use NewBytes(&HEF, &HBF, &HBD), &HFFFD&
            Test.Use NewBytes(&HF4, &H8F, &HBF, &HBF), &HDFFFDBFF ' "U-0010FFFF"

        Case "T3_Malformed_1_UnexpectedContinuation_3"
            Test.Use NewBytes(&H80)
            Test.Use NewBytes(&HBF)
            Test.Use NewBytes(&H80, &HBF)
            Test.Use NewBytes(&H80, &HBF, &H80)
            Test.Use NewBytes(&H80, &HBF, &H80, &HBF)
            Test.Use NewBytes(&H80, &HBF, &H80, &HBF, &H80)
            Test.Use NewBytes(&H80, &HBF, &H80, &HBF, &H80, &HBF)
            Test.Use NewBytes(&H80, &HBF, &H80, &HBF, &H80, &HBF, &H80)
            Test.Use NewBytes(&H80, &H81, &H82, &H83, &H84, &H85, &H86, &H87, &H88, &H89, &H8A, &H8B, &H8C, &H8D, &H8E, &H8F, _
                &H90, &H91, &H92, &H93, &H94, &H95, &H96, &H97, &H98, &H99, &H9A, &H9B, &H9C, &H9D, &H9E, &H9F, _
                &HA0, &HA1, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7, &HA8, &HA9, &HAA, &HAB, &HAC, &HAD, &HAE, &HAF, _
                &HB0, &HB1, &HB2, &HB3, &HB4, &HB5, &HB6, &HB7, &HB8, &HB9, &HBA, &HBB, &HBC, &HBD, &HBE, &HBF)
                
        Case "T3_Malformed_2_LonelyStart_3"
            Test.Use NewBytes(&HC0, &H20, &HC1, &H20, &HC2, &H20, &HC3, &H20, &HC4, &H20, &HC5, &H20, &HC6, &H20, &HC7, &H20, _
                &HC8, &H20, &HC9, &H20, &HCA, &H20, &HCB, &H20, &HCC, &H20, &HCD, &H20, &HCE, &H20, &HCF, &H20, _
                &HD0, &H20, &HD1, &H20, &HD2, &H20, &HD3, &H20, &HD4, &H20, &HD5, &H20, &HD6, &H20, &HD7, &H20, _
                &HD8, &H20, &HD9, &H20, &HDA, &H20, &HDB, &H20, &HDC, &H20, &HDD, &H20, &HDE, &H20, &HDF, &H20)
            Test.Use NewBytes(&HE0, &H20, &HE1, &H20, &HE2, &H20, &HE3, &H20, &HE4, &H20, &HE5, &H20, &HE6, &H20, &HE7, &H20, _
                &HE8, &H20, &HE9, &H20, &HEA, &H20, &HEB, &H20, &HEC, &H20, &HED, &H20, &HEE, &H20, &HEF, &H20)
            Test.Use NewBytes(&HF0, &H20, &HF1, &H20, &HF2, &H20, &HF3, &H20, &HF4, &H20, &HF5, &H20, &HF6, &H20, &HF7, &H20)
            Test.Use NewBytes(&HFC, &H20, &HFD, &H20)
            
        Case "T3_Malformed_3_LastContinuationMissing"
            Test.Use NewBytes(&HC0)
            Test.Use NewBytes(&HE0, &H80)
            Test.Use NewBytes(&HF0, &H80, &H80)
            Test.Use NewBytes(&HF8, &H80, &H80, &H80)
            Test.Use NewBytes(&HFC, &H80, &H80, &H80, &H80)
            Test.Use NewBytes(&HF7, &HBF, &HBF)
            Test.Use NewBytes(&HF, &HBF, &HBF, &HBF)
            Test.Use NewBytes(&HFD, &HBF, &HBF, &HBF, &HBF)
            Test.Use NewBytes(&HDF)
            Test.Use NewBytes(&HEF, &HBF)
            
        Case "T3_Malformed_5_ImpossibleBytes"
            Test.Use NewBytes(&HFE)
            Test.Use NewBytes(&HFF)
            Test.Use NewBytes(&HFE, &HFE, &HFF, &HFF)
        
        Case "T4_Overlong_1_ASCII_Slash"
            Test.Use NewBytes(&HC0, &HAF)
            Test.Use NewBytes(&HE0, &H80, &HAF)
            Test.Use NewBytes(&HF8, &H80, &H80, &H80, &HAF)
            Test.Use NewBytes(&HFC, &H80, &H80, &H80, &H80, &HAF)
        
        Case "T4_Overlong_2_MaximumBoundary"
            Test.Use NewBytes(&HC1, &HBF)
            Test.Use NewBytes(&HE0, &H9F, &HBF)
            Test.Use NewBytes(&HF0, &H8F, &HBF, &HBF)
            Test.Use NewBytes(&HF8, &H87, &HBF, &HBF, &HBF)
            Test.Use NewBytes(&HFC, &H83, &HBF, &HBF, &HBF, &HBF)
        
        Case "T4_Overlong_3_NUL"
            Test.Use NewBytes(&HC0, &H80)
            Test.Use NewBytes(&HE0, &H80, &H80)
            Test.Use NewBytes(&HF0, &H80, &H80, &H80)
            Test.Use NewBytes(&HF8, &H80, &H80, &H80, &H80)
            Test.Use NewBytes(&HFC, &H80, &H80, &H80, &H80, &H80)
            
        Case "T5_IllegalCodePosition_1_UTF16Surrogates"
            Test.Use NewBytes(&HED, &HA0, &H80)
            Test.Use NewBytes(&HED, &HAD, &HBF)
            Test.Use NewBytes(&HED, &HAE, &H80)
            Test.Use NewBytes(&HED, &HAF, &HBF)
            Test.Use NewBytes(&HED, &HB0, &H80)
            Test.Use NewBytes(&HED, &HBE, &H80)
            Test.Use NewBytes(&HED, &HBF, &HBF)
        
        Case "T5_IllegalCodePosition_2_PairedUTF16Surrogates"
            Test.Use NewBytes(&HED, &HA0, &H80, &HED, &HB0, &H80)
            Test.Use NewBytes(&HED, &HA0, &H80, &HED, &HBF, &HBF)
            Test.Use NewBytes(&HED, &HAD, &HBF, &HED, &HB0, &H80)
            Test.Use NewBytes(&HED, &HAD, &HBF, &HED, &HBF, &HBF)
            Test.Use NewBytes(&HED, &HAE, &H80, &HED, &HB0, &H80)
            Test.Use NewBytes(&HED, &HAE, &H80, &HED, &HBF, &H8F)
            Test.Use NewBytes(&HED, &HAF, &HBF, &HED, &HB0, &H80)
            Test.Use NewBytes(&HED, &HAF, &HBF, &HED, &HBF, &HBF)
        
        Case "T5_IllegalCodePosition_3_Other"
            Test.Use NewBytes(&HEF, &HBF, &HBE), 65534
            Test.Use NewBytes(&HEF, &HBF, &HBF), 65535
        
        Case "GetByteCount_WithValues_ReturnsExpected"
            Test.Use "", 0
            Test.Use "abc", 3
            Test.Use t("\x00"), 1
            Test.Use t("\x7f"), 1
            Test.Use t("\x80"), 2
            Test.Use t("\x7ff"), 2
            Test.Use t("\x800"), 3
            Test.Use t("\xffff"), 3
            Test.Use t("\xd800\xdc00"), 4
            Test.Use t("\xdbff\xdfff"), 4
            Test.Use t("\xd800\x80"), 5
            Test.Use t("a\xd800"), 4
            Test.Use t("\xd800a"), 4
            Test.Use t("\xd800\xd800\xdc00"), 7
            
        Case "GetByteCount_WithFailingFallback_ReturnsExpected"
            Test.Use t("\xd800"), 0
            Test.Use t("\xdc00"), 0
            Test.Use t("\xd800a"), 1
            Test.Use t("a\xd800"), 1
            Test.Use t("a\xd800a"), 2
            Test.Use t("a\xdc00a"), 2
            Test.Use t("\xd800\xd800\xdc00"), 4
            
        Case "GetByteCount_WhenUsingExceptionFallbackWithValues_ThrowsEncoderFallbackException"
            Test.Use "\u00C4 \uD802\u0033 \u00AE", 2, &HD802
            Test.Use "\u00C4 \uDC02\u0033 \u00AE", 2, &HDC02
            Test.Use " \uD802", 1, &HD802
            Test.Use " \uDC02", 1, &HDC02
            Test.Use "a\uDC02a", 1, &HDC02
            Test.Use "a\uD802a", 1, &HD802
            
        Case "GetCharCount_WithValues_ReturnsExpected"
            Test.Use NewBytes(), 0
            Test.Use NewBytes(&H41), 1 ' A
            Test.Use NewBytes(&H41, &H42, &H43), 3 ' ABC
            Test.Use NewBytes(&HC0), 1
            Test.Use NewBytes(&HC2, &H80), 1  ' U+0080
            Test.Use NewBytes(&HC0, &H80), 2 ' 2-byte encoding of U+0000 is invalid
            Test.Use NewBytes(223, 191), 1 ' U+07FF
            Test.Use NewBytes(224, 160, 128), 1 ' U+0800
            Test.Use NewBytes(239, 191, 191), 1 ' U+FFFF
            Test.Use NewBytes(224, 159, 191), 2 ' U+07FF
            Test.Use NewBytes(239, 191, 189), 1 ' U+FFFD
            Test.Use NewBytes(&HED, &HA0, &H80), 2 ' U+D800
            Test.Use NewBytes(240, 144, 128, 128), 2 ' U+D800 U+DC00
            Test.Use NewBytes(240, 65), 2
            Test.Use NewBytes(&H80), 1
            Test.Use NewBytes(&HF6, &HB0, &H80, &H80), 4 ' U+D800 U+0000
            
        Case "GetCharsEx_WithByteSequence_DecodesCharacters"
            ' Single byte sequence
            Test.Use NewBytes(), NewIntegers()
            Test.Use NewBytes(&H41), NewChars("A") ' A
            Test.Use NewBytes(&H41, &H42, &H43), NewChars("A", "B", "C") ' ABC
            
            ' Two-byte sequence
            Test.Use NewBytes(&HC2, &H80), NewChars(&H80)  ' U+0080
            Test.Use NewBytes(223, 191), NewChars(&H7FF) ' U+07FF
            
            ' Three-byte sequence
            Test.Use NewBytes(224, 160, 128), NewChars(&H800) ' U+0800
            Test.Use NewBytes(&HED, &H9F, &HBF), NewChars(&HD7FF)
            Test.Use NewBytes(&HEE, &H80, &H80), NewChars(&HE000)
            Test.Use NewBytes(239, 191, 191), NewChars(&HFFFF) ' U+FFFF
            
            ' Four-byte sequence (surrogate pairs)
            Test.Use NewBytes(240, 144, 128, 128), NewChars(&HD800, &HDC00) ' U+D800 U+DC00
            
            ' Test multi-character byte sequences
            Test.Use NewBytes(&H41, &HC2, &H81, 224, 160, 128, &HEE, &H80, &H80), NewChars("A", &H81, &H800, &HE000) ' A U+0081 U+0800 U+E000
            
            ' Invalid sequences
            Test.Use NewBytes(&HC0), NewChars(&HFFFD)
            Test.Use NewBytes(&HC0, &H80), NewChars(&HFFFD, &HFFFD) ' 2-byte encoding of U+0000 is invalid
            Test.Use NewBytes(224, 159, 191), NewChars(&HFFFD, &HFFFD) ' U+07FF -> U+FFFD U+FFFD
            Test.Use NewBytes(239, 191, 189), NewChars(&HFFFD) ' U+FFFD
            Test.Use NewBytes(&HED, &HA0, &H80), NewChars(&HFFFD, &HFFFD) ' U+D800
            Test.Use NewBytes(240, 65), NewChars(&HFFFD, &H41)
            Test.Use NewBytes(&H80), NewChars(&HFFFD)
            Test.Use NewBytes(&HF6, &HB0, &H80, &H80), NewChars(&HFFFD, &HFFFD, &HFFFD, &HFFFD) ' U+D800 U+0000
            Test.Use NewBytes(&HF0, &H10, &HFC, &H0), NewChars(&HFFFD, &H10, &HFFFD, &H0)
            
        Case "GetBytesEx_WithInvalidSurrogate_ThrowsEncoderFallbackException"
            Test.Use "\u00C4 \uD802\u0033 \u00AE", 2, &HD802
            Test.Use "\u00C4 \uDC02\u0033 \u00AE", 2, &HDC02
            Test.Use " \uD802", 1, &HD802
            Test.Use " \uDC02", 1, &HDC02
            
        Case "GetBytesEx_WithCharacters_EncodesCharacters"
            Test.Use NewChars(0, 127, 1, 126, 2, 125), NewBytes(0, 127, 1, 126, 2, 125)
            Test.Use NewChars(&H80, &H7FF, &H81, &H7FE), NewBytes(&HC2, &H80, &HDF, &HBF, &HC2, &H81, &HDF, &HBE)
            Test.Use NewChars(&H800, &HFFFF, &H801, &HFFFE), NewBytes(&HE0, &HA0, &H80, &HEF, &HBF, &HBF, &HE0, &HA0, &H81, &HEF, &HBF, &HBE)
            Test.Use NewChars(&HD800, &HDC00, &HDBFF, &HDFFF), NewBytes(&HF0, &H90, &H80, &H80, &HF4, &H8F, &HBF, &HBF)
            Test.Use NewChars(&HDC00, 97), NewBytes(&HEF, &HBF, &HBD, 97)
            Test.Use NewChars(&HD800, 97), NewBytes(&HEF, &HBF, &HBD, 97)
            Test.Use NewChars(97, &HD800), NewBytes(97, &HEF, &HBF, &HBD)
            Test.Use NewChars(), NewBytes()
            
        Case "GetCharsEx_WithInvalidByteSequence_ThrowsDecoderFallbackException"
            Test.Use NewBytes(&HC0), 0, NewBytes(&HC0)
            Test.Use NewBytes(1, 2, &H80), 2, NewBytes(&H80) ' invalid start of any sequence
            Test.Use NewBytes(1, 2, &HC1), 2, NewBytes(&HC1) ' invalid start of 2-byte sequence
            Test.Use NewBytes(1, 2, &HC2, 0), 2, NewBytes(&HC2) ' invalid 2nd byte of 2-byte character
            Test.Use NewBytes(&HF0, &H10, &HFC, &H0), 0, NewBytes(&HF0)
            Test.Use NewBytes(&HF0, &H10, &HF8, &H0), 0, NewBytes(&HF0)
            Test.Use NewBytes(1, 2, &HE1), 2, NewBytes(&HE1) ' hanging 1 byte of 3-byte character
            Test.Use NewBytes(1, 2, &HE0, &HA2), 2, NewBytes(&HE0, &HA2) ' hanging 2 bytes of 3-byte character
            Test.Use NewBytes(&HF0, &H8F, &HBF, &HBF), 0, NewBytes(&HF0, &H8F)
            Test.Use NewBytes(1, 2, &HF0, &H90, &H80), 2, NewBytes(&HF0, &H90, &H80) ' hanging first 3 bytes of surrogate pair
            Test.Use NewBytes(1, 2, &HF0, &H90), 2, NewBytes(&HF0, &H90) ' hanging 2 bytes of surrogate pair

        Case "GetCharCount_WithInvalidByteSequence_ThrowsDecoderFallbackException"
            Test.Use NewBytes(&HC0), 0, NewBytes(&HC0)
            Test.Use NewBytes(1, 2, &H80), 2, NewBytes(&H80) ' invalid start of any sequence
            Test.Use NewBytes(1, 2, &HC1), 2, NewBytes(&HC1) ' invalid start of 2-byte sequence
            Test.Use NewBytes(1, 2, &HC2, 0), 2, NewBytes(&HC2) ' invalid 2nd byte of 2-byte character
            Test.Use NewBytes(&HF0, &H10, &HFC, &H0), 0, NewBytes(&HF0)
            Test.Use NewBytes(&HF0, &H10, &HF8, &H0), 0, NewBytes(&HF0)
            Test.Use NewBytes(1, 2, &HE1), 2, NewBytes(&HE1) ' hanging start of 3-byte character
            Test.Use NewBytes(1, 2, &HE0, &HA2), 2, NewBytes(&HE0, &HA2) ' hanging 2 bytes of 3-byte character
            Test.Use NewBytes(1, 2, &HF0, &H90, &H80), 2, NewBytes(&HF0, &H90, &H80) ' hanging first 3 bytes of surrogate pair
            Test.Use NewBytes(&HF0, &H8F, &HBF, &HBF), 0, NewBytes(&HF0, &H8F)
            Test.Use NewBytes(1, 2, &HF0, &H90), 2, NewBytes(&HF0, &H90) ' hanging 2 bytes of surrogate pair
        
    End Select
End Sub

Public Sub Setup()
    Set mUTF8 = NewUTF8Encoding(True, True)
End Sub

Public Sub GetCharCount_WithValues_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected)
    Dim Enc As New UTF8Encoding
    
    Dim Actual As Long
    Actual = Enc.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharsEx_WithByteSequence_DecodesCharacters(ByRef Bytes() As Byte, ByRef ExpectedChars() As Integer)
    Dim Enc             As New UTF8Encoding
    Dim ExpectedCount   As Long
    Dim ActualCount     As Long
    Dim ActualChars()   As Integer
    
    ExpectedCount = CorArray.Length(ExpectedChars)
    ActualChars = CorArray.CreateInstance(vbInteger, ExpectedCount)
    
    ActualCount = Enc.GetCharsEx(Bytes, 0, CorArray.Length(Bytes), ActualChars, 0)
    
    Assert.That ActualCount, Iz.EqualTo(ExpectedCount), "Count"
    Assert.That ActualChars, Iz.EqualTo(ExpectedChars), "Chars"
End Sub

' This test is to help test the fast loop with only ASCII characters. It's only
' useful in debug mode so the code can be walked through to ensure expected
' code paths are being executed.
Public Sub GetCharsEx_WithLongSequence_DecodesChars()
    Const Count As Long = 21
    Dim Enc As New UTF8Encoding
    Dim Bytes() As Byte
    Dim Chars() As Integer
    Dim i As Long
    
    ReDim Bytes(0 To Count - 1)
    ReDim Chars(0 To Count - 1)
    
    For i = 0 To Count - 1
        Bytes(i) = i
    Next
    
    Assert.That Enc.GetCharsEx(Bytes, 0, Count, Chars, 0), Iz.EqualTo(Count)
    
    Assert.That Chars, Iz.EqualTo(Bytes)
End Sub

Public Sub GetCharsEx_WithInvalidByteSequence_ThrowsDecoderFallbackException(ByRef InvalidSequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Enc As UTF8Encoding
    Dim Chars(100) As Integer
    
    Set Enc = NewUTF8Encoding(False, True)
    
    On Error GoTo Catch
    Enc.GetCharsEx InvalidSequence, LBound(InvalidSequence), CorArray.Length(InvalidSequence), Chars, 0
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    
    ' The .NET version of the fallback will have a DecoderFallbackException.Index of one less than
    ' what the documentation says. The documentation says the index is of the byte that caused
    ' the invalid sequence to happen. I'm going with the documentation because that index value
    ' makes the most sense, so all VBCorLib indexes will be one more than what .NET indicates.
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetCharsEx_WithInvalidByteSequenceAndNonZeroLBound_ThrowsDecoderFallbackException()
    Dim Enc As UTF8Encoding
    Dim Chars(100) As Integer
    Dim Bytes(5 To 10) As Byte
    
    Set Enc = NewUTF8Encoding(False, True)
    Bytes(7) = &H80
    
    On Error GoTo Catch
    Enc.GetCharsEx Bytes, 5, 6, Chars, 0
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(7)
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&H80))
End Sub

Public Sub GetCharCount_WithInvalidByteSequence_ThrowsDecoderFallbackException(ByRef InvalidSequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Enc As UTF8Encoding
    
    Set Enc = NewUTF8Encoding(False, True)
    
    On Error GoTo Catch
    Enc.GetCharCount InvalidSequence
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetByteCount_WithFailingFallback_ReturnsExpected(ByRef s As String, ByVal Expected As Long)
    Dim Enc As UTF8Encoding
    Set Enc = Encoding.UTF8.Clone
    Set Enc.EncoderFallback = New NullEncoderFallback
    
    Dim Actual As Long
    Actual = Enc.GetByteCount(s)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WhenUsingExceptionFallbackWithValues_ThrowsEncoderFallbackException(ByRef Chars As String, ByVal ExpectedIndex As Long, ByVal ExpectedChar As Integer)
    Dim Enc As Encoding
    Set Enc = NewUTF8Encoding(False, True)
    
    On Error Resume Next
    Enc.GetByteCount t(Chars)
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.CharUnknown, Iz.EqualTo(ExpectedChar), "CharUnknown"
End Sub

Public Sub GetByteCount_WithValues_ReturnsExpected(ByRef s As String, ByVal Expected As Long)
    Dim Enc As New UTF8Encoding
    
    Dim Actual As Long
    Actual = Enc.GetByteCount(s)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsBrowserDisplay_WhenCalled_ReturnsTrue()
    Assert.That mUTF8.IsBrowserDisplay, Iz.True
End Sub

Public Sub IsBrowserSave_WhenCalled_ReturnsTrue()
    Assert.That mUTF8.IsBrowserSave, Iz.True
End Sub

Public Sub IsMailNewsDisplay_WhenCalled_ReturnsTrue()
    Assert.That mUTF8.IsMailNewsDisplay, Iz.True
End Sub

Public Sub IsMailNewsSave_WhenCalled_ReturnsTrue()
    Assert.That mUTF8.IsMailNewsSave, Iz.True
End Sub

Public Sub GetBytes_WithString_ReturnsEncoded()
    Dim UTF8Enc As New UTF8Encoding
    
    Dim Actual() As Byte
    Actual = UTF8Enc.GetBytes(t("\u0041\u2262\u0391\u002E"))
    
    Dim Expected() As Byte
    Expected = NewBytes(&H41, &HE2, &H89, &HA2, &HCE, &H91, &H2E)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithTwoHighSurrogates_ThrowsEncoderFallbackException()
    Dim Encoding As UTF8Encoding
    Set Encoding = NewUTF8Encoding(False, True)
    
    On Error Resume Next
    Encoding.GetBytes t("\uD800\uD800")
    
    AssertEncoderFallbackException Err
End Sub

Public Sub GetBytesEx_WithInvalidSurrogate_ThrowsEncoderFallbackException(ByRef InvalidSurrogate As String, ByVal ExpectedIndex As Long, ByVal ExpectedChar As Integer)
    Dim Enc As Encoding
    Dim Invalid As String
    Dim Bytes() As Byte
    
    Set Enc = NewUTF8Encoding(False, True)
    Invalid = t(InvalidSurrogate)
    ReDim Bytes(0 To Enc.GetMaxByteCount(Len(Invalid)))
    
    On Error Resume Next
    Enc.GetBytesEx Invalid, 0, Len(Invalid), Bytes, 0
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.CharUnknown, Iz.EqualTo(ExpectedChar), "CharUnknown"
End Sub

Public Sub GetBytesEx_WithCharacters_EncodesCharacters(ByRef Chars As Variant, ByRef ExpectedBytes() As Byte)
    Dim Enc As New UTF8Encoding
    Dim ExpectedCount As Long
    Dim ActualCount As Long
    Dim ActualBytes() As Byte
    
    ExpectedCount = CorArray.Length(ExpectedBytes)
    ActualBytes = CorArray.CreateInstance(vbByte, ExpectedCount)
    
    ActualCount = Enc.GetBytesEx(Chars, 0, CorArray.Length(Chars), ActualBytes, 0)
    
    Assert.That ActualCount, Iz.EqualTo(ExpectedCount), "Count"
    Assert.That ActualBytes, Iz.EqualTo(ExpectedBytes), "Bytes"
End Sub

Public Sub GetBytesEx_WithSmallConversionBuffer_ThrowsArgumentException()
    Dim Enc As New UTF8Encoding
    Dim Bytes() As Byte
    Dim Chars() As Integer
    
    Chars = NewChars("a")
    Bytes = NewBytes()
    
    On Error Resume Next
    Enc.GetBytesEx Chars, 0, 1, Bytes, 0
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub GetBytes_WithChars_ReturnsEncoded()
    Dim UTF8Enc As New UTF8Encoding
    Dim Chars() As Integer
    Chars = CorString.ToCharArray(t("\u0048\u0069\u0020\u004D\u006F\u006D\u0020\u263A\u0021"))
    
    Dim Actual(0 To 10) As Byte
    Dim ByteCount       As Long
    ByteCount = UTF8Enc.GetBytesEx(Chars, 0, CorArray.Length(Chars), Actual, 0)
    
    Dim Expected() As Byte
    Expected = NewBytes(&H48, &H69, &H20, &H4D, &H6F, &H6D, &H20, &HE2, &H98, &HBA, &H21)
    Assert.That ByteCount, Iz.EqualTo(11)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WhenCalled_ReturnsDecoded()
    Dim UTF8Enc As New UTF8Encoding
    Dim UTF8Bytes() As Byte
    UTF8Bytes = NewBytes(&H41, &HE2, &H89, &HA2, &HCE, &H91, &H2E)
    
    Dim Actual() As Integer
    Actual = UTF8Enc.GetChars(UTF8Bytes)
    
    Dim Expected() As Integer
    Expected = NewIntegers(&H41, &H2262, &H391, &H2E)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub MaxCharCount_With50AndNoBOM_Returns51()
    Dim UTF8Enc As New UTF8Encoding
    
    Dim Actual As Long
    Actual = UTF8Enc.GetMaxCharCount(50)
    
    Assert.That Actual, Iz.EqualTo(51)
End Sub

Public Sub MaxByteCount_With50AndBOM_Returns153()
    Dim UTF8Enc As UTF8Encoding
    Set UTF8Enc = Cor.NewUTF8Encoding(True)
    
    Dim Actual As Long
    Actual = UTF8Enc.GetMaxByteCount(50)
    
    Assert.That Actual, Iz.EqualTo(153)
End Sub

Public Sub MaxByteCount_With50AndNoBOM_Returns153()
    Dim UTF8Enc As New UTF8Encoding
    
    Dim Actual As Long
    Actual = UTF8Enc.GetMaxByteCount(50)
    
    Assert.That Actual, Iz.EqualTo(153)
End Sub

Public Sub RegressionTest_Bug_59648()
    Dim UTF8Enc As UTF8Encoding
    Dim Data() As Byte
    Dim s As String
    Dim Chars() As Integer
    
    Set UTF8Enc = Cor.NewUTF8Encoding(True)
    Data = NewBytes(&HC0, &HAF)
    Assert.That UTF8Enc.GetCharCount(Data), Iz.EqualTo(2)
    
    s = UTF8Enc.GetString(Data)
    Assert.That s, Iz.EqualTo(t("\uFFFD\uFFFD"))

    Data = NewBytes(&H30, &H31, &HC0, &HAF, &H30, &H32)
    Chars = CorString.ToCharArray(UTF8Enc.GetString(Data))

    Dim Expected() As Integer
    Expected = NewIntegers(&H30, &H31, &HFFFD, &HFFFD, &H30, &H32)
    Assert.That Chars, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharsEx_Correct_GreekWord_kosme()
    Dim Enc             As New UTF8Encoding
    Dim Data()          As Byte
    Dim ActualChars(4)  As Integer
    Dim ActualCount     As Long
    
    Data = NewBytes(&HCE, &HBA, &HE1, &HBD, &HB9, &HCF, &H83, &HCE, &HBC, &HCE, &HB5)
    
    ActualCount = Enc.GetCharsEx(Data, 0, UBound(Data) + 1, ActualChars, 0)
    
    Dim Expected() As Integer
    Expected = NewChars(&H3BA, &H1F79, &H3C3, &H3BC, &H3B5)
    Assert.That ActualCount, Iz.EqualTo(5)
    Assert.That ActualChars, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_Correct_GreekWord_kosme()
    Dim Enc     As New UTF8Encoding
    Dim Data()  As Byte
    Dim Actual  As Long
    
    Data = NewBytes(&HCE, &HBA, &HE1, &HBD, &HB9, &HCF, &H83, &HCE, &HBC, &HCE, &HB5)
    
    Actual = Enc.GetCharCount(Data)
    
    Assert.That Actual, Iz.EqualTo(5)
End Sub

Public Sub T2_Boundary_1_FirstPossibleSequence_Pass(ByRef Data() As Byte, ByVal Expected As Long)
    Dim s As String
        
    s = mUTF8.GetString(Data)
        
    Assert.That MemLong(StrPtr(s)), Iz.EqualTo(Expected)
    Assert.That mUTF8.GetBytes(s), Iz.EqualTo(Data), "GetBytes"
End Sub

Public Sub GetCharsEx_Boundary_1_FirstPossibleSequence_Fail_5()
    Dim Data215() As Byte
    Dim Chars(0 To 100) As Integer
    Data215 = NewBytes(&HF8, &H88, &H80, &H80, &H80)
    
    On Error GoTo Catch
    mUTF8.GetCharsEx Data215, 0, CorArray.Length(Data215), Chars, 0
    
Catch:
    AssertDecoderFallbackException Err
End Sub

Public Sub GetCharCount_Boundary_1_FirstPossibleSequence_Fail_5()
    Dim Data215() As Byte
    Data215 = NewBytes(&HF8, &H88, &H80, &H80, &H80)
    
    On Error GoTo Catch
    mUTF8.GetCharCount Data215
    
Catch:
    AssertDecoderFallbackException Err
End Sub

Public Sub GetCharsEx_Boundary_2_FirstPossibleSequence__Fail_6()
    Dim Data() As Byte
    Dim Chars(10) As Integer
    Data = NewBytes(&HFC, &H84, &H80, &H80, &H80, &H80)
    
    On Error GoTo Catch
    mUTF8.GetCharsEx Data, 0, UBound(Data) + 1, Chars, 0
    
Catch:
    AssertDecoderFallbackException Err
End Sub

Public Sub GetCharCount_Boundary_2_FirstPossibleSequence__Fail_6()
    Dim Data() As Byte
    Data = NewBytes(&HFC, &H84, &H80, &H80, &H80, &H80)
    
    On Error GoTo Catch
    mUTF8.GetCharCount Data
    
Catch:
    AssertDecoderFallbackException Err
End Sub

Public Sub T2_Boundary_2_LastPossibleSequence_Pass(ByRef Data() As Byte, ByVal Expected As Long)
    Dim s As String
    
    s = mUTF8.GetString(Data)
    
    Assert.That MemLong(StrPtr(s)), Iz.EqualTo(Expected)
    Assert.That mUTF8.GetBytes(s), Iz.EqualTo(Data)
End Sub

Public Sub GetCharsEx_T2_Boundary_2_LastPossibleSequence_Fail_4()
    Dim Data() As Byte
    Dim Chars(100) As Integer
    
    Data = NewBytes(&H7F, &HBF, &HBF, &HBF)
    
    On Error GoTo Catch
    mUTF8.GetCharsEx Data, 0, Len1D(Data), Chars, 0
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(1), "Count"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&HBF)), "BytesUnknown"
End Sub

Public Sub GetCharCount_T2_Boundary_2_LastPossibleSequence_Fail_4()
    Dim Data() As Byte
    
    Data = NewBytes(&H7F, &HBF, &HBF, &HBF)
    
    On Error GoTo Catch
    mUTF8.GetCharCount Data
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(1), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&HBF)), "BytesUnknown"
End Sub

Public Sub GetCharsEx_T2_Boundary_2_LastPossibleSequence_Fail_5()
    Dim Data() As Byte
    Dim Chars(100) As Integer
    
    Data = NewBytes(&HFB, &HBF, &HBF, &HBF, &HBF)
    
    On Error GoTo Catch
    mUTF8.GetCharsEx Data, 0, Len1D(Data), Chars, 0
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(0), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&HFB))
End Sub

Public Sub GetCharCount_T2_Boundary_2_LastPossibleSequence_Fail_5()
    Dim Data() As Byte
    
    Data = NewBytes(&HFB, &HBF, &HBF, &HBF, &HBF)
    
    On Error GoTo Catch
    mUTF8.GetCharCount Data
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(0), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&HFB))
End Sub

Public Sub GetCharsEx_T2_Boundary_2_LastPossibleSequence_Fail_6()
    Dim Data() As Byte
    Dim Chars(100) As Integer
    
    Data = NewBytes(&HFB, &HBF, &HBF, &HBF, &HBF, &HBF)
    
    On Error GoTo Catch
    mUTF8.GetCharsEx Data, 0, Len1D(Data), Chars, 0
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(0), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&HFB))
End Sub

Public Sub GetCharCount_T2_Boundary_2_LastPossibleSequence_Fail_6()
    Dim Data() As Byte
    
    Data = NewBytes(&HFB, &HBF, &HBF, &HBF, &HBF, &HBF)
    
    On Error GoTo Catch
    mUTF8.GetCharCount Data
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(0), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&HFB))
End Sub

Public Sub T2_Boundary_3_Other_Pass(ByRef Data() As Byte, ByVal Expected As Long)
    Dim s As String
    
    s = mUTF8.GetString(Data)
    
    Assert.That MemLong(StrPtr(s)), Iz.EqualTo(Expected), "Chars"
    Assert.That mUTF8.GetBytes(s), Iz.EqualTo(Data), "GetBytes"
End Sub

Public Sub GetCharsEx_T2_Boundary_3_Other_Fail_5()
    Dim Data() As Byte
    Dim Chars(100) As Integer
    
    Data = NewBytes(&HF4, &H90, &H80, &H80)
    
    On Error GoTo Catch
    mUTF8.GetCharsEx Data, 0, Len1D(Data), Chars, 0
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(0), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&HF4, &H90))
End Sub

Public Sub GetCharCount_T2_Boundary_3_Other_Fail_5()
    Dim Data() As Byte
    
    Data = NewBytes(&HF4, &H90, &H80, &H80)
    
    On Error GoTo Catch
    mUTF8.GetCharCount Data
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(0), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&HF4, &H90))
End Sub

Public Sub T3_Malformed_1_UnexpectedContinuation_3(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertDecoderFallbackException Err
End Sub

Public Sub T3_Malformed_2_LonelyStart_3(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertDecoderFallbackException Err
End Sub

Public Sub T3_Malformed_3_LastContinuationMissing(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertDecoderFallbackException Err
End Sub

Public Sub GetCharsEx_T3_Malformed_4_ConcatenationImcomplete()
    Dim Data() As Byte
    Dim Chars(100) As Integer
    
    Data = NewBytes(&HC0, &HE0, &H80, &HF0, &H80, &H80, &HF8, &H80, &H80, &H80, &HFC, &H80, &H80, &H80, &H80, &HDF, _
                    &HEF, &HBF, &HF7, &HBF, &HBF, &HFB, &HBF, &HBF, &HBF, &HFD, &HBF, &HBF, &HBF, &HBF)
        
    On Error GoTo Catch
    mUTF8.GetCharsEx Data, 0, Len1D(Data), Chars, 0
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(0), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&HC0))
End Sub

Public Sub GetCharCount_T3_Malformed_4_ConcatenationImcomplete()
    Dim Data() As Byte
    
    Data = NewBytes(&HC0, &HE0, &H80, &HF0, &H80, &H80, &HF8, &H80, &H80, &H80, &HFC, &H80, &H80, &H80, &H80, &HDF, _
                    &HEF, &HBF, &HF7, &HBF, &HBF, &HFB, &HBF, &HBF, &HBF, &HFD, &HBF, &HBF, &HBF, &HBF)
        
    On Error GoTo Catch
    mUTF8.GetCharCount Data
    
Catch:
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(0), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(NewBytes(&HC0))
End Sub

Public Sub T3_Malformed_5_ImpossibleBytes(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertDecoderFallbackException Err
End Sub

Public Sub T4_Overlong_1_ASCII_Slash(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertDecoderFallbackException Err
End Sub

Public Sub T4_Overlong_2_MaximumBoundary(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertDecoderFallbackException Err
End Sub

Public Sub T4_Overlong_3_NUL(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertDecoderFallbackException Err
End Sub

Public Sub T5_IllegalCodePosition_1_UTF16Surrogates(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertDecoderFallbackException Err
End Sub

Public Sub T5_IllegalCodePosition_2_PairedUTF16Surrogates(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertDecoderFallbackException Err
End Sub

' MS Fx 1.1 passes this
Public Sub T5_IllegalCodePosition_3_Other(ByRef Data() As Byte, ByVal Expected As Long)
    Dim Actual As String
    
    Actual = mUTF8.GetString(Data)
    
    Assert.That MemLong(StrPtr(Actual)), Iz.EqualTo(Expected)
End Sub

Public Sub GetCharsFEFF()
    Dim Data() As Byte
    Dim Enc As UTF8Encoding
    Dim s As String
    
    Data = NewBytes(&HEF, &HBB, &HBF)
    Set Enc = NewUTF8Encoding(False, True)
    s = Enc.GetString(Data)
    Assert.That s, Iz.EqualTo(t("\uFEFF"))
    
    Set Enc = Encoding.UTF8
    Dim TestChars() As Integer
    Dim Bytes() As Byte
    Dim Chars() As Integer
    TestChars = NewChars(t("\uFEFF"), "A")
    Bytes = Enc.GetBytes(TestChars)
    Chars = Enc.GetChars(Bytes)
    
    Assert.That Chars(0), Iz.EqualTo(&HFEFF)
    Assert.That Chars(1), Iz.EqualTo(Asc("A"))
End Sub

Public Sub Bug77315()
    Dim Enc As Encoding
    Set Enc = NewUTF8Encoding(False, True)
    
    On Error Resume Next
    Enc.GetString NewBytes(&HED, &HA2, &H8C)
    
    AssertDecoderFallbackException Err
End Sub

Public Sub SufficientByteArray2()
    Dim Bytes(10) As Byte
    Dim u As Encoding
    Set u = Encoding.UTF8
    
    Assert.That u.GetByteCount(t("\ufffd")), Iz.EqualTo(3), "#1-1"
    Assert.That u.GetByteCount(t("\uD800")), Iz.EqualTo(3), "#1-2"
    Assert.That u.GetByteCount(t("\uDC00")), Iz.EqualTo(3), "#1-3"
    Assert.That u.GetByteCount(t("\uD800\uDC00")), Iz.EqualTo(4), "#1-4"
    Assert.That u.GetBytesEx(t("\uDC00"), 0, 1, Bytes, 0), Iz.EqualTo(3), "#1-5"
    
    Assert.That CorArray.Length(u.GetBytes(t("\uFFFD"))), Iz.EqualTo(3), "#2-1"
    Assert.That CorArray.Length(u.GetBytes(t("\uD800"))), Iz.EqualTo(3), "#2-2"
    Assert.That CorArray.Length(u.GetBytes(t("\uDC00"))), Iz.EqualTo(3), "#2-3"
    Assert.That CorArray.Length(u.GetBytes(t("\uD800\uDC00"))), Iz.EqualTo(4), "#2-4"
    
    On Error Resume Next
    NewUTF8Encoding(False, True).GetBytes NewIntegers(&HDF45, &HD808), 0, 2
    AssertEncoderFallbackException Err
End Sub

Public Sub DecoderFallbackSimple()
    Dim e As Encoding
    Dim Chars(100) As Integer
    
    Set e = NewUTF8Encoding(False, False)
    Assert.That e.GetDecoder.GetCharCount(NewBytes(183), 0, 1), Iz.EqualTo(1), "#1"
    Assert.That e.GetDecoder.GetChars(NewBytes(183), 0, 1, Chars, 0), Iz.EqualTo(1), "#2"
    Assert.That e.GetString(NewBytes(183)), Iz.EqualTo(t("\ufffd")), "#3"
End Sub

Public Sub Serialization_WhenCalled_DoesNotDuplicateExceptionFallbacks()
    Dim Original    As UTF8Encoding
    Dim Copy        As UTF8Encoding
    Set Original = Cor.NewUTF8Encoding(True, True)
    
    Set Copy = Persist(Original)
    
    Assert.That Copy.DecoderFallback, Iz.SameAs(DecoderFallback.ExceptionFallback)
    Assert.That Copy.EncoderFallback, Iz.SameAs(EncoderFallback.ExceptionFallback)
End Sub







