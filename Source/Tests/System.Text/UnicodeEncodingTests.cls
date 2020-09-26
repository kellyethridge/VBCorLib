VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnicodeEncodingTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UnicodeEncodingTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource


Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetCharCount_WithDefaultFallback_ReturnsExpected"
            Test.Use NewBytes(), 0
            Test.Use NewBytes(0), 1
            Test.Use NewBytes(0, 0), 1
            Test.Use NewBytes(0, 0, 0), 2
            Test.Use NewBytes(0, 0, 0, 0), 2
            Test.Use NewBytes(&HFF), 1
            Test.Use NewBytes(&HFF, &HFF), 1
            Test.Use NewBytes(&HFF, &HFF, &HFF), 2
            Test.Use NewBytes(&HFF, &HFF, &HFF, &HFF), 2
            Test.Use NewBytes(&HD8), 1
            Test.Use NewBytes(&HFF, &HD8), 1
            Test.Use NewBytes(&H0, &HD8), 1
            Test.Use NewBytes(&H40, &HFF, &HD8), 2
            Test.Use NewBytes(&HFF, &HD8, &HFF, &HFF), 2
            Test.Use NewBytes(&H40, &HD8, &HFF, &HD8, &HFF, &H40), 3
            Test.Use NewBytes(&H40, &HDC, &HFF, &HD8, &HFF, &H40), 3
            Test.Use NewBytes(&HD8, &H0, &HDC, &H0), 2
            
        Case "GetCharCount_WithEmptyFallback_ReturnsExpected"
            Test.Use NewBytes(), 0
            Test.Use NewBytes(0), 0
            Test.Use NewBytes(0, 0), 1
            Test.Use NewBytes(0, 0, 0), 1
            Test.Use NewBytes(0, 0, 0, 0), 2
            Test.Use NewBytes(&HFF), 0
            Test.Use NewBytes(&HFF, &HFF), 1
            Test.Use NewBytes(&HFF, &HFF, &HFF), 1
            Test.Use NewBytes(&HFF, &HFF, &HFF, &HFF), 2
            Test.Use NewBytes(&HD8), 0
            Test.Use NewBytes(&HFF, &HD8), 0
            Test.Use NewBytes(&H0, &HD8), 0
            Test.Use NewBytes(&H40, &HFF, &HD8), 1
            Test.Use NewBytes(&HFF, &HD8, &HFF, &HFF), 1
            Test.Use NewBytes(&H40, &HD8, &HFF, &HD8, &HFF, &H40), 1
            Test.Use NewBytes(&H40, &HDC, &HFF, &HD8, &HFF, &H40), 1
            Test.Use NewBytes(&HD8, &H0, &HDC, &H0), 2
            
        Case "GetCharCount_WithInvalidByteSequence_ThrowsDecoderFallbackException"
            Test.Use NewBytes(&H40), 0, NewBytes(&H40)
            Test.Use NewBytes(&H40, &H0, &H0, &HD8, &H40, &H0), 4, NewBytes(&H0, &HD8)
            Test.Use NewBytes(&H40, &H0, &H0, &HD8, &H1, &HD8), 4, NewBytes(&H0, &HD8)
            Test.Use NewBytes(&H40, &H0, &H0, &HDC, &H1, &HD8), 2, NewBytes(&H0, &HDC)
            Test.Use NewBytes(&H40, &H0, &H1, &HD8), 2, NewBytes(&H1, &HD8)
            
        Case "GetCharCount_WithInvalidByteSequenceAndBigEndian_ThrowsDecoderFallbackException"
            Test.Use NewBytes(&H40), 0, NewBytes(&H40)
            Test.Use NewBytes(&H0, &H40, &HD8, &H1), 2, NewBytes(&HD8, &H1)
                                
        Case "GetChars_WithValues_ReturnsExpected"
            Test.Use NewBytes(), NewIntegers()
            Test.Use NewBytes(0), NewIntegers(&HFFFD)
            Test.Use NewBytes(0, 0), NewIntegers(0)
            Test.Use NewBytes(0, 0, 0), NewIntegers(0, &HFFFD)
            Test.Use NewBytes(&H0, &HD8), NewIntegers(&HFFFD)
            Test.Use NewBytes(&H0, &HD8, &H0, &HD8), NewIntegers(&HFFFD, &HFFFD)
            Test.Use NewBytes(&H0, &HD8, &H0, &H0), NewIntegers(&HFFFD, &H0)
            Test.Use NewBytes(&H0, &HDC), NewIntegers(&HFFFD)
            Test.Use NewBytes(&H0, &HDC, &H0, &HDC), NewIntegers(&HFFFD, &HFFFD)
            Test.Use NewBytes(&HA0, &H3, &HA3, &H3), NewIntegers(&H3A0, &H3A3)
            Test.Use NewBytes(&H0, &HDC, &H0, &H0), NewIntegers(&HFFFD, &H0)
            
        Case "GetCharsEx_WithValues_Decodes"
             Test.Use NewBytes(&H0, &HD8, &H41, &H0), 2, NewChars(&HFFFD, &H41)
             Test.Use NewBytes(&H41, &H0, &H0, &HD8, &H0, &HDC, &H41, &H0, &H42, &H0, &H43, &H0), 6, NewChars(&H41, &HD800, &HDC00, &H41, &H42, &H43)

        Case "GetCharsEx_WithInvalidSequence_ThrowsDecoderFallbackException"
            Test.Use NewBytes(&H40), 0, NewBytes(&H40)
            Test.Use NewBytes(0, &H40, &H41), 2, NewBytes(&H41)
            Test.Use NewBytes(0, &H40, &H0, &HD8, &H0, &H41), 4, NewBytes(&H0, &HD8)
            Test.Use NewBytes(0, &H40, &H0, &HD8, &H1, &HD8), 4, NewBytes(&H0, &HD8)
            Test.Use NewBytes(0, &H40, &H0, &HDC, &H1, &HD8), 2, NewBytes(&H0, &HDC)
            Test.Use NewBytes(0, &H40, &H0, &HD8), 2, NewBytes(&H0, &HD8)
            
        Case "GetCharsEx_WhenBigEndianWithInvalidSequence_ThrowsDecoderFallbackException"
            Test.Use NewBytes(&H40), 0, NewBytes(&H40)
            Test.Use NewBytes(&H40, 0, &H41), 2, NewBytes(&H41)
            Test.Use NewBytes(&H40, 0, &HD8, 0, &H41, 0), 4, NewBytes(&HD8, 0)
        
        Case "GetByteCount_WithValues_ReturnsExpected"
            Test.Use t("za"), 4
            Test.Use t("\xd800\xdc00"), 4
            Test.Use t("\xd800a"), 4
            Test.Use t("\xd800\xd800"), 4
            Test.Use t("\xd800\xd800a"), 6
            Test.Use t("A\xdc00"), 4
            Test.Use t("\uD800\uD802\uDC00"), 6
            
        Case "GetByteCount_WhenUsingReplacementFallbackWithValues_ReturnsExpected"
            Test.Use t("za"), 4
            Test.Use t("\xd800\xdc00"), 4
            Test.Use t("\xd800a"), 2
            Test.Use t("\xd800\xd800"), 0
            Test.Use t("\xd800\xd800a"), 2
            
        Case "GetByteCount_WhenUsingExceptionFallbackWithValues_ThrowsEncoderFallbackException"
            Test.Use "\u00C4 \uD802\u0033 \u00AE", 2, &HD802
            Test.Use "\u00C4 \uDC02\u0033 \u00AE", 2, &HDC02
            Test.Use " \uD802", 1, &HD802
            Test.Use " \uDC02", 1, &HDC02
            Test.Use "a\uDC02a", 1, &HDC02
            Test.Use "a\uD802a", 1, &HD802
            
        Case "GetBytesEx_WithLittleEndianEncoding_ReturnsExpected"
            Test.Use t("za"), NewBytes(122, 0, 97, 0)
            Test.Use t("\xd800\xdc00"), NewBytes(&H0, &HD8, &H0, &HDC)
            Test.Use t("\xd800a"), NewBytes(&HFD, &HFF, 97, 0)
            Test.Use t("\xd800\xd800"), NewBytes(&HFD, &HFF, &HFD, &HFF)
            Test.Use t("\xd800\xd800a"), NewBytes(&HFD, &HFF, &HFD, &HFF, 97, 0)
        
        Case "GetBytes_WithLittleEndianEncoding_ReturnsExpected"
            Test.Use t("\xD800\xD800"), NewBytes(&HFD, &HFF, &HFD, &HFF)
            Test.Use t("\u03a0\u03a3"), NewBytes(&HA0, &H3, &HA3, &H3)
            Test.Use t("\xD800\xDC00"), NewBytes(&H0, &HD8, &H0, &HDC)
            Test.Use t("A\xD800"), NewBytes(&H41, &H0, &HFD, &HFF)
            Test.Use t("A\xDC00"), NewBytes(&H41, &H0, &HFD, &HFF)
            
        Case "GetBytes_WithBigEndianEncoding_ReturnsEncoded"
            Test.Use t("\xD800\xD800"), NewBytes(&HFF, &HFD, &HFF, &HFD)
            Test.Use t("\u03a0\u03a3"), NewBytes(&H3, &HA0, &H3, &HA3)
            Test.Use t("\xD800\xDC00"), NewBytes(&HD8, &H0, &HDC, &H0)
            Test.Use t("A\xD800"), NewBytes(&H0, &H41, &HFF, &HFD)
            Test.Use t("A\xDC00"), NewBytes(&H0, &H41, &HFF, &HFD)
            
        Case "GetBytesEx_WithInvalidSurrogates_IncludesFallbackCharacters"
            Test.Use t("\uD8FF"), NewBytes(&HFD, &HFF)
            Test.Use t("\uD8FF\uD8FF"), NewBytes(&HFD, &HFF, &HFD, &HFF)
            Test.Use t("\uD8FF\uDC00\uD8FF"), NewBytes(&HFF, &HD8, &H0, &HDC, &HFD, &HFF)
            Test.Use t("\uDC00\uD8FF\uDC00\uD8FF"), NewBytes(&HFD, &HFF, &HFF, &HD8, &H0, &HDC, &HFD, &HFF)
            Test.Use t("\uD8FFa"), NewBytes(&HFD, &HFF, &H61, &H0)
        
        Case "GetBytesEx_WithInvalidSurrogate_ThrowsEncoderFallbackException"
            Test.Use "\u00C4 \uD802\u0033 \u00AE", 2, &HD802
            Test.Use "\u00C4 \uDC02\u0033 \u00AE", 2, &HDC02
            Test.Use " \uD802", 1, &HD802
            Test.Use " \uDC02", 1, &HDC02
            
    End Select
End Sub

Public Sub IsBrowserDisplay_WhenCalled_ReturnsFalse()
    Dim Enc As New UnicodeEncoding
    Assert.That Enc.IsBrowserDisplay, Iz.False
End Sub

Public Sub IsBrowserSave_WhenCalled_ReturnsTrue()
    Dim Enc As New UnicodeEncoding
    Assert.That Enc.IsBrowserSave, Iz.True
End Sub

Public Sub IsMailNewsDisplay_WhenCalled_ReturnsFalse()
    Dim Enc As New UnicodeEncoding
    Assert.That Enc.IsMailNewsDisplay, Iz.False
End Sub

Public Sub IsMailNewsSave_WhenCalled_ReturnsFalse()
    Dim Enc As New UnicodeEncoding
    Assert.That Enc.IsMailNewsSave, Iz.False
End Sub

Public Sub GetBytexEx_WhenLittleEndian_Encodes()
    Dim Enc As New UnicodeEncoding
    Dim Unicode As String
    Unicode = t("\u03a0\u03a3")
    
    Dim ActualCount As Long
    Dim ActualBytes(3) As Byte
    ActualCount = Enc.GetBytesEx(Unicode, 0, Len(Unicode), ActualBytes, 0)
    
    Assert.That ActualCount, Iz.EqualTo(4)
    Assert.That ActualBytes, Iz.EqualTo(NewBytes(&HA0, &H3, &HA3, &H3))
End Sub

Public Sub GetBytesEx_WithSubCharArray_EncodesSubArray()
    Dim Enc As New UnicodeEncoding
    Dim Unicode As String
    Dim Bytes() As Byte
    
    Unicode = "abcdefghijk"
    ReDim Bytes(0 To 8)
    
    Assert.That Enc.GetBytesEx(Unicode, 3, 3, Bytes, 1), Iz.EqualTo(6)
    
    Assert.That Bytes, Iz.EqualTo(NewBytes(0, 100, 0, 101, 0, 102, 0, 0, 0))
End Sub

Public Sub GetBytes_WithLittleEndianEncoding_ReturnsExpected(ByRef Chars As Variant, ByRef Expected() As Byte)
    Dim Enc As New UnicodeEncoding
    
    Dim Actual() As Byte
    Actual = Enc.GetBytes(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytesEx_WithLittleEndianEncoding_ReturnsExpected(ByRef Chars As String, ByRef Expected() As Byte)
    Dim Enc As New UnicodeEncoding
    Dim Actual() As Byte
    Dim Count As Long
    
    ReDim Actual(UBound(Expected))
    
    Count = Enc.GetBytesEx(Chars, 0, Len(Chars), Actual, 0)
    
    Assert.That Count, Iz.EqualTo(CorArray.Length(Expected)), "#Count"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithBigEndianEncoding_ReturnsEncoded(ByRef Chars As Variant, ByRef Expected() As Byte)
    Dim Enc As UnicodeEncoding
    Set Enc = NewUnicodeEncoding(True, True)
    
    Dim Actual() As Byte
    Actual = Enc.GetBytes(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithLargeCharacterSet_ReturnsEncoded()
    Dim Chars() As Integer
    Dim Actual() As Byte
    Dim Expected() As Byte
    Dim i As Long
    Dim s As String
    
    ReDim Chars(0 To 55)
    
    For i = 0 To 25
        Chars(i) = 65 + i
        Chars(30 + i) = 97 + i
    Next
    
    Chars(26) = &HD800
    Chars(27) = &HDC00
    Chars(28) = &HDBFF
    Chars(29) = &HDFFF
    s = NewString(Chars)
    Expected = s
    
    Actual = Encoding.Unicode.GetBytes(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytesEx_WithInvalidSurrogates_IncludesFallbackCharacters(ByRef InvalidSurrogates As Variant, ByRef Expected() As Byte)
    Dim Enc As New UnicodeEncoding
    Dim Bytes() As Byte
    Dim Length As Long
    
    ReDim Bytes(0 To 100)
    Length = Enc.GetBytesEx(InvalidSurrogates, 0, Len(InvalidSurrogates), Bytes, 0)
    
    ReDim Preserve Bytes(0 To Length - 1)
    Assert.That Bytes, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytesEx_WithInvalidSurrogate_ThrowsEncoderFallbackException(ByVal Invalid As String, ByVal ExpectedIndex As Long, ByVal ExpectedChar As Integer)
    Dim Enc As Encoding
    Dim Bytes() As Byte
    
    Set Enc = NewUnicodeEncoding(False, False, True)
    Invalid = t(Invalid)
    ReDim Bytes(0 To Enc.GetMaxByteCount(Len(Invalid)))
    
    On Error Resume Next
    Enc.GetBytesEx Invalid, 0, Len(Invalid), Bytes, 0
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.CharUnknown, Iz.EqualTo(ExpectedChar), "CharUnknown"
End Sub

Public Sub GetBytesEx_TestOptimizationUseOfSmallerBytesBuffer_ThrowsArgumentException()
    Dim Chars() As Integer
    Dim Bytes() As Byte
    
    Chars = NewChars(66, 67, 68, 69, 70)
    ReDim Bytes(0 To 8)
    
    On Error Resume Next
    Encoding.Unicode.GetBytesEx Chars, 0, 5, Bytes, 0
    
    AssertArgumentException Err
End Sub

' This test is useful only when stepping the the UnicodeEncoding.GetBytesCore method.
Public Sub GetBytesEx_WithMisalignedBytesArray_DoesNotUseOptimizedRoutine()
    Dim Chars() As Integer
    Dim Bytes() As Byte
    
    Chars = NewChars(66, 67, 68, 69, 70)
    ReDim Bytes(0 To 10)
    
    Assert.That Encoding.Unicode.GetBytesEx(Chars, 0, 5, Bytes, 1), Iz.EqualTo(10)
    
    Assert.That Bytes, Iz.EqualTo(NewBytes(0, 66, 0, 67, 0, 68, 0, 69, 0, 70, 0))
End Sub

' This test is useful only when stepping the the UnicodeEncoding.GetCharsCore method.
Public Sub GetCharsEx_WithMisalignedBytesArray_DoesNotUseOptimizedRoutine()
    Dim Chars() As Integer
    Dim Bytes() As Byte

    Bytes = NewBytes(0, 66, 0, 67, 0, 68, 0, 69, 0, 70, 0)
    ReDim Chars(0 To 4)
    
    Assert.That Encoding.Unicode.GetCharsEx(Bytes, 1, 10, Chars, 0), Iz.EqualTo(5)
    
    Assert.That Chars, Iz.EqualTo(NewChars(66, 67, 68, 69, 70))
End Sub

Public Sub GetChars_WithValues_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim Enc As New UnicodeEncoding
    Dim Actual() As Integer
    
    Actual = Enc.GetChars(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharsEx_WithValues_Decodes(ByRef Bytes() As Byte, ByVal ExpectedCount As Long, ByRef ExpectedChars() As Integer)
    Dim Enc As New UnicodeEncoding
    Dim ActualCount As Long
    Dim ActualChars() As Integer
    
    ReDim ActualChars(UBound(ExpectedChars))
    
    ActualCount = Enc.GetCharsEx(Bytes, 0, CorArray.Length(Bytes), ActualChars, 0)
    
    Assert.That ActualCount, Iz.EqualTo(ExpectedCount), "Count"
    Assert.That ActualChars, Iz.EqualTo(ExpectedChars), "Chars"
End Sub

Public Sub GetCharsEx_WithSmallOutputBuffer_ThrowsArgumentException()
    Dim Enc As New UnicodeEncoding
    Dim SmallOutput(0) As Integer
    Dim Bytes(0 To 3) As Byte
    
    On Error Resume Next
    Enc.GetCharsEx Bytes, 0, 4, SmallOutput, 0
    
    AssertArgumentException Err, "Chars"
End Sub

Public Sub GetCharsEx_WithInvalidSequence_ThrowsDecoderFallbackException(ByRef InvalidSequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Enc As UnicodeEncoding
    Dim Chars() As Integer
    
    Set Enc = NewUnicodeEncoding(False, False, True)
    ReDim Chars(0 To Enc.GetMaxCharCount(CorArray.Length(InvalidSequence)))
    
    On Error Resume Next
    Enc.GetCharsEx InvalidSequence, LBound(InvalidSequence), CorArray.Length(InvalidSequence), Chars, 0
    
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetCharsEx_WhenBigEndianWithInvalidSequence_ThrowsDecoderFallbackException(ByRef InvalidSequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Enc As UnicodeEncoding
    Dim Chars() As Integer
    
    Set Enc = NewUnicodeEncoding(True, False, True)
    ReDim Chars(0 To Enc.GetMaxCharCount(CorArray.Length(InvalidSequence)))
    
    On Error Resume Next
    Enc.GetCharsEx InvalidSequence, LBound(InvalidSequence), CorArray.Length(InvalidSequence), Chars, 0
    
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetCharCount_WhenCalled_ReturnsExpectedCount()
    Dim b() As Byte: b = NewBytes(255, 254, 115, 0, 104, 0, 105, 0)
    Dim Enc As New UnicodeEncoding
    
    Dim Actual As Long
    Actual = Enc.GetCharCount(b, 2, CorArray.Length(b) - 2)
    
    Assert.That Actual, Iz.EqualTo(3)
End Sub

Public Sub GetCharCount_WithInvalidByteSequence_ThrowsDecoderFallbackException(ByRef Sequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Enc As Encoding
    
    Set Enc = NewUnicodeEncoding(False, False, True)
    
    On Error Resume Next
    Enc.GetCharCount Sequence
    
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Ex)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetCharCount_WithInvalidByteSequenceAndBigEndian_ThrowsDecoderFallbackException(ByRef Sequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Enc As Encoding
    
    Set Enc = NewUnicodeEncoding(True, False, True)
    
    On Error Resume Next
    Enc.GetCharCount Sequence
    
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Ex)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetPreamble_WhenLittleEndian_ReturnsPreambleBytes()
    Dim Enc As UnicodeEncoding: Set Enc = NewUnicodeEncoding(False, True)
    
    Dim Actual() As Byte
    Actual = Enc.GetPreamble
    
    Assert.That Actual, Iz.EqualTo(NewBytes(&HFF, &HFE))
End Sub

Public Sub GetPreamble_WhenBigEndian_ReturnsPreambleBytes()
    Dim Enc As UnicodeEncoding: Set Enc = NewUnicodeEncoding(True, True)
    
    Dim Actual() As Byte
    Actual = Enc.GetPreamble
    
    Assert.That Actual, Iz.EqualTo(NewBytes(&HFE, &HFF))
End Sub

Public Sub GetPreamble_WhenLittleEndianAndNoBOM_ReturnsEmpty()
    Dim Enc As UnicodeEncoding: Set Enc = NewUnicodeEncoding(False, False)
    
    Dim Actual() As Byte
    Actual = Enc.GetPreamble
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub GetMaxCharCount_WhenCalled_ReturnsExpected()
    Dim Enc As New UnicodeEncoding
    
    Assert.That Enc.GetMaxCharCount(50), Iz.EqualTo(26)
    Assert.That Enc.GetMaxCharCount(51), Iz.EqualTo(27)
End Sub

Public Sub GetMaxByteCount_WhenCalled_ReturnsExpected()
    Dim Enc As New UnicodeEncoding
    
    Assert.That Enc.GetMaxByteCount(50), Iz.EqualTo(102)
End Sub

Public Sub ZeroLengthArrays()
    Dim Enc As New UnicodeEncoding
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

Public Sub ByteOrderMark()
    Dim LittleEndianString As String: LittleEndianString = t("\ufeff\u0042\u004f\u004d")
    Dim BigEndianString As String: BigEndianString = t("\ufffe\u4200\u4f00\u4d00")
    Dim LittleEndianBytes() As Byte: LittleEndianBytes = NewBytes(&HFF, &HFE, &H42, &H0, &H4F, &H0, &H4D, &H0)
    Dim BigEndianBytes() As Byte: BigEndianBytes = NewBytes(&HFE, &HFF, &H0, &H42, &H0, &H4F, &H0, &H4D)
    Dim Enc As UnicodeEncoding
    
    Set Enc = Cor.NewUnicodeEncoding(False, True)
    Assert.That Enc.GetBytes(LittleEndianString), Iz.EqualTo(LittleEndianBytes), "BOM #1"
    Assert.That Enc.GetBytes(BigEndianString), Iz.EqualTo(BigEndianBytes), "BOM #2"
    Assert.That Enc.GetString(LittleEndianBytes), Iz.EqualTo(LittleEndianString), "BOM #3"
    Assert.That Enc.GetString(BigEndianBytes), Iz.EqualTo(BigEndianString), "BOM #4"
    
    Set Enc = Cor.NewUnicodeEncoding(True, True)
    Assert.That Enc.GetBytes(LittleEndianString), Iz.EqualTo(BigEndianBytes), "BOM #5"
    Assert.That Enc.GetBytes(BigEndianString), Iz.EqualTo(LittleEndianBytes), "BOM #6"
    Assert.That Enc.GetString(LittleEndianBytes), Iz.EqualTo(BigEndianString), "BOM #7"
    Assert.That Enc.GetString(BigEndianBytes), Iz.EqualTo(LittleEndianString), "BOM #8"
End Sub

Public Sub GetString_WithOddCountOf0x00_ReturnsExpected()
    Dim Actual As String

    Actual = Encoding.Unicode.GetString(NewBytes(0, 0, 0), 0, 3)

    Assert.That Left$(Actual, 1), Iz.EqualTo(vbNullChar), "0"
    Assert.That Len(Actual), Iz.EqualTo(2), "Length"
    Assert.That Mid$(Actual, 2, 1), Iz.EqualTo(ChrW$(&HFFFD)), "1"
End Sub

Public Sub GetString_WithOddCountOf0xFF_ReturnsExpected()
    Dim Actual As String

    Actual = Encoding.Unicode.GetString(NewBytes(&HFF, &HFF, &HFF), 0, 3)

    Assert.That Left$(Actual, 1), Iz.EqualTo(ChrW$(&HFFFF)), "0"
    Assert.That Len(Actual), Iz.EqualTo(2), "Length"
    Assert.That Mid$(Actual, 2, 1), Iz.EqualTo(ChrW$(&HFFFD)), "1"
End Sub

Public Sub GetCharCount_WithEmptyFallback_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Enc As UnicodeEncoding
    Set Enc = Encoding.Unicode.Clone
    Set Enc.DecoderFallback = Cor.NewDecoderReplacementFallback("")
    
    Dim Actual As Long
    Actual = Enc.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithDefaultFallback_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = Encoding.Unicode.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithValues_ReturnsExpected(ByRef Chars As Variant, ByVal Expected As Long)
    Dim Enc As New UnicodeEncoding
    Dim Actual As Long
    
    Actual = Enc.GetByteCount(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WhenUsingExceptionFallbackWithValues_ThrowsEncoderFallbackException(ByRef Chars As String, ByVal ExpectedIndex As Long, ByVal ExpectedChar As Integer)
    Dim Enc As UnicodeEncoding
    Set Enc = NewUnicodeEncoding(False, False, True)
    
    On Error Resume Next
    Enc.GetByteCount t(Chars)
    
    Dim Ex As EncoderFallbackException
    Set Ex = AssertEncoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.CharUnknown, Iz.EqualTo(ExpectedChar), "CharUnknown"
End Sub

Public Sub GetByteCount_WhenUsingReplacementFallbackWithValues_ReturnsExpected(ByRef Chars As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = MakeEncoding(Replacement:="").GetByteCount(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithSubsetOfArray_ReturnsExpectedCount()
    Dim Actual As Long
    
    Actual = MakeEncoding.GetByteCount("abcd", 1, 2)
    
    Assert.That Actual, Iz.EqualTo(4)
End Sub

Public Sub Serialization_WhenCalled_DoesNotDuplicateExceptionFallbacks()
    Dim Original    As UnicodeEncoding
    Dim Copy        As UnicodeEncoding
    Set Original = Cor.NewUnicodeEncoding(False, False, True)
    
    Set Copy = Persist(Original)
    
    Assert.That Copy.DecoderFallback, Iz.SameAs(DecoderFallback.ExceptionFallback)
    Assert.That Copy.EncoderFallback, Iz.SameAs(EncoderFallback.ExceptionFallback)
End Sub

Private Function MakeEncoding(Optional ByVal BigEndian As Boolean, Optional ByVal Replacement As String = "?") As UnicodeEncoding
    Set MakeEncoding = Cor.NewUnicodeEncoding(BigEndian, True).Clone
    Set MakeEncoding.EncoderFallback = Cor.NewEncoderReplacementFallback(Replacement)
End Function











