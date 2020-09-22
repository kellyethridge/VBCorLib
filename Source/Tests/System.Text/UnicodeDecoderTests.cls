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
        Case "GetCharCount_WithSingleCall_ReturnsExpected"
            Test.Use NewBytes(), 0
            Test.Use NewBytes(0), 0
            Test.Use NewBytes(0, 0), 1
            Test.Use NewBytes(0, 0, 0), 1
            Test.Use NewBytes(0, 0, 0, 0), 2
            Test.Use NewBytes(&H0, &HD8), 0
            Test.Use NewBytes(&H0, &HDC), 1
            Test.Use NewBytes(&H0, &HD8, &H0), 0
            Test.Use NewBytes(&H0, &HD8, &H0, &HDC), 2
            Test.Use NewBytes(&H0, &HD8, &H0, &H41, &H0, &H42), 3
            
        Case "GetCharCount_WithMultipleCalls_DoesNotStoreState"
            Test.Use NewBytes(), NewBytes(), 0
            Test.Use NewBytes(0), NewBytes(), 0
            Test.Use NewBytes(), NewBytes(0), 0
            Test.Use NewBytes(0, &HD8), NewBytes(), 0
            Test.Use NewBytes(0, &HD8), NewBytes(0), 0
            Test.Use NewBytes(0, &HD8), NewBytes(0, &HDC), 1
            Test.Use NewBytes(0, &H41), NewBytes(), 1
            Test.Use NewBytes(0, &H41), NewBytes(0, &H41), 2
            
        Case "GetCharCount_WithHangingByteSequenceAndFlushing_ThrowsDecoderFallbackException"
            Test.Use NewBytes(0, &H40, 0), 2, NewBytes(0)
            Test.Use NewBytes(0, &H40, 0, &HD8), 2, NewBytes(0, &HD8)
            Test.Use NewBytes(0, &H40, &H41), 2, NewBytes(&H41)
            
        Case "GetChars_WithValues_ReturnsExpected"
            Test.Use NewBytes(), NewIntegers()
            Test.Use NewBytes(0), NewIntegers()
            Test.Use NewBytes(&H41, &H0), NewIntegers(&H41)
            Test.Use NewBytes(0, 0, 0), NewIntegers(0)
            Test.Use NewBytes(0, 0, 0, 0), NewIntegers(0, 0)
            Test.Use NewBytes(0, &HD8), NewIntegers()
            Test.Use NewBytes(&H0, &HD8, &H0), NewIntegers()
            Test.Use NewBytes(&H0, &HD8, &H0, &HDC), NewIntegers(&HD800, &HDC00)
            Test.Use NewBytes(&H0, &HD8, &H0, &H41, &H0, &H42), NewIntegers(&HFFFD, &H4100, &H4200)
        
        Case "GetChars_WithMultipleCallsWithoutFlush_ReturnsExpected"
            Test.Use NewBytes(), NewBytes(), NewIntegers(), NewIntegers(), 0
            Test.Use NewBytes(0), NewBytes(), NewIntegers(), NewIntegers(), 0
            Test.Use NewBytes(), NewBytes(0), NewIntegers(), NewIntegers(), 0
            Test.Use NewBytes(0), NewBytes(0), NewIntegers(), NewIntegers(0), 1
            Test.Use NewBytes(0, &HD8, 0, &HDC), NewBytes(), NewIntegers(&HD800, &HDC00), NewIntegers(), 2
            Test.Use NewBytes(0, &HD8), NewBytes(0, &HDC), NewIntegers(), NewIntegers(&HD800, &HDC00), 2
            Test.Use NewBytes(&H41, 0, 0, &HD8), NewBytes(0, &HDC), NewIntegers(&H41), NewIntegers(&HD800, &HDC00), 3
        
        Case "GetChars_WithMultipleCallsWithFlush_ReturnsExpected"
            Test.Use NewBytes(), NewBytes(), NewIntegers(), NewIntegers(), 0
            Test.Use NewBytes(0), NewBytes(), NewIntegers(), NewIntegers(&HFFFD), 1
            Test.Use NewBytes(), NewBytes(0), NewIntegers(), NewIntegers(&HFFFD), 1
            Test.Use NewBytes(0), NewBytes(0), NewIntegers(), NewIntegers(0), 1
            Test.Use NewBytes(0, &HD8, 0, &HDC), NewBytes(), NewIntegers(&HD800, &HDC00), NewIntegers(), 2
            Test.Use NewBytes(0, &HD8), NewBytes(0, &HDC), NewIntegers(), NewIntegers(&HD800, &HDC00), 2
            Test.Use NewBytes(&H41, 0, 0, &HD8), NewBytes(0, &HDC), NewIntegers(&H41), NewIntegers(&HD800, &HDC00), 3
            Test.Use NewBytes(), NewBytes(0, &HD8), NewIntegers(), NewIntegers(&HFFFD), 1
            
        Case "GetChars_WithInvalidSequence_ThrowsDecoderFallbackException"
            Test.Use NewBytes(0, &H40, &H41), 2, NewBytes(&H41)
            Test.Use NewBytes(0, &H40, &H0, &HD8), 2, NewBytes(&H0, &HD8)
            
    End Select
End Sub

Public Sub GetCharCount_WithSingleCall_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(Bytes, 0, CorArray.Length(Bytes))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WithMultipleCalls_DoesNotStoreState(ByRef b1() As Byte, ByRef b2() As Byte, ByVal Expected As Long)
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(b1, 0, CorArray.Length(b1))
    Actual = Actual + Decoder.GetCharCount(b2, 0, CorArray.Length(b2))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharCount_WhenFlushingWithHangingHighSurrogate_UsesFallback()
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(NewBytes(&H0, &H41, &H0, &HD8), 0, 4, True)
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub GetCharCount_WithHangingByteSequenceAndFlushing_ThrowsDecoderFallbackException(ByRef Sequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Decoder As Decoder
    Set Decoder = NewUnicodeEncoding(False, False, True).GetDecoder
    
    On Error Resume Next
    Decoder.GetCharCount Sequence, 0, CorArray.Length(Sequence), True
    
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex)
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown)
End Sub

Public Sub GetCharCount_WithOddBytes_DoesNotUseFallback()
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(NewBytes(&H0, &H41, &H41), 0, 3)
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub GetCharCount_WithFlushingWithOddBytes_UsesFallback()
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(NewBytes(&H0, &H41, &H41), 0, 3, True)
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub GetCharCount_WithSurrogatePairFollowedByCharacter_ReturnsTwo()
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(NewBytes(&H0, &HD8), 0, 2)
    Actual = Actual + Decoder.GetCharCount(NewBytes(&H0, &HDC), 0, 2)
    Actual = Actual + Decoder.GetCharCount(NewBytes(&H0, &H41), 0, 2)
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub GetCharCount_WithOddBytes_ReturnsOne()
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Actual = Decoder.GetCharCount(NewBytes(&H0), 0, 1)
    Actual = Actual + Decoder.GetCharCount(NewBytes(&H0, &H0, &H0), 0, 3)
    Actual = Actual + Decoder.GetCharCount(NewBytes(&H0), 0, 1)
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub GetChars_WithValues_ReturnsExpected(ByRef Bytes() As Byte, ByRef ExpectedChars() As Integer)
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Dim Chars() As Integer
    Chars = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(Bytes, 0, CorArray.Length(Bytes)))
    Actual = Decoder.GetChars(Bytes, 0, CorArray.Length(Bytes), Chars, 0)
    
    Assert.That Actual, Iz.EqualTo(CorArray.Length(Chars))
    Assert.That Chars, Iz.EqualTo(ExpectedChars)
End Sub

Public Sub GetChars_WithMultipleCallsWithoutFlush_ReturnsExpected(ByRef b1() As Byte, ByRef b2() As Byte, ByRef ExpectedChars1() As Integer, ByRef ExpectedChars2() As Integer, ByVal ExpectedLength As Long)
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Dim Chars1() As Integer
    Dim Chars2() As Integer
    Chars1 = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(b1, 0, CorArray.Length(b1)))
    Actual = Decoder.GetChars(b1, 0, CorArray.Length(b1), Chars1, 0)
    Chars2 = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(b2, 0, CorArray.Length(b2)))
    Actual = Actual + Decoder.GetChars(b2, 0, CorArray.Length(b2), Chars2, 0)
    
    Assert.That Actual, Iz.EqualTo(ExpectedLength)
    Assert.That Chars1, Iz.EqualTo(ExpectedChars1)
    Assert.That Chars2, Iz.EqualTo(ExpectedChars2)
End Sub

Public Sub GetChars_WithMultipleCallsWithFlush_ReturnsExpected(ByRef b1() As Byte, ByRef b2() As Byte, ByRef ExpectedChars1() As Integer, ByRef ExpectedChars2() As Integer, ByVal ExpectedLength As Long)
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As Long
    Dim Chars1() As Integer
    Dim Chars2() As Integer
    Chars1 = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(b1, 0, CorArray.Length(b1)))
    Actual = Decoder.GetChars(b1, 0, CorArray.Length(b1), Chars1, 0)
    Chars2 = CorArray.CreateInstance(vbInteger, Decoder.GetCharCount(b2, 0, CorArray.Length(b2), True))
    Actual = Actual + Decoder.GetChars(b2, 0, CorArray.Length(b2), Chars2, 0, True)
    
    Assert.That Actual, Iz.EqualTo(ExpectedLength)
    Assert.That Chars1, Iz.EqualTo(ExpectedChars1)
    Assert.That Chars2, Iz.EqualTo(ExpectedChars2)
End Sub

Public Sub GetChars_WithInvalidSequence_ThrowsDecoderFallbackException(ByRef Sequence() As Byte, ByVal ExpectedIndex As Long, ByRef ExpectedBytesUnknown() As Byte)
    Dim Decoder As Decoder
    Dim Chars(100) As Integer
    
    Set Decoder = NewUnicodeEncoding(False, False, True).GetDecoder
    
    On Error Resume Next
    Decoder.GetChars Sequence, 0, CorArray.Length(Sequence), Chars, 0, True
    
    Dim Ex As DecoderFallbackException
    Set Ex = AssertDecoderFallbackException(Err)
    Assert.That Ex.Index, Iz.EqualTo(ExpectedIndex), "Index"
    Assert.That Ex.BytesUnknown, Iz.EqualTo(ExpectedBytesUnknown), "BytesUnknown"
End Sub

Public Sub GetChars_WithLastByteSaved_ClearsLastByte()
    Dim Decoder As Decoder
    Dim Chars(100) As Integer
    
    Set Decoder = NewUnicodeEncoding(False, False, True).GetDecoder
    
    Assert.That Decoder.GetChars(NewBytes(0), 0, 1, Chars, 0, False), Iz.EqualTo(0)
    Assert.That Decoder.GetChars(NewBytes(&H40), 0, 1, Chars, 0, True), Iz.EqualTo(1)
    Assert.That Decoder.GetChars(NewBytes(&H0, &H41), 0, 2, Chars, 0, True), Iz.EqualTo(1)
End Sub

Public Sub GetChars_WithLastCharSaved_ClearsLastChar()
    Dim Decoder As Decoder
    Dim Chars(100) As Integer
    
    Set Decoder = NewUnicodeEncoding(False, False, True).GetDecoder

    Assert.That Decoder.GetChars(NewBytes(0, &HD8), 0, 2, Chars, 0, False), Iz.EqualTo(0)
    Assert.That Decoder.GetChars(NewBytes(0, &HDC), 0, 2, Chars, 0, True), Iz.EqualTo(2)
    Assert.That Decoder.GetChars(NewBytes(0, &H40), 0, 2, Chars, 0, True), Iz.EqualTo(1)
End Sub

Public Sub Fallback_WhenCalled_ReturnsSameInstanceAsEncoding()
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim Actual As DecoderFallback
    Set Actual = Decoder.Fallback
    
    Assert.That Actual, Iz.SameAs(Encoding.Unicode.DecoderFallback)
End Sub

Public Sub Fallback_WhenSetToNothing_ThrowsArgumentNullException()
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    On Error Resume Next
    Set Decoder.Fallback = Nothing
    
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub FallbackBuffer_WhenCalled_ReturnsCachedInstance()
    Dim Decoder As Decoder
    Set Decoder = Encoding.Unicode.GetDecoder
    
    Dim b1 As DecoderFallbackBuffer
    Dim b2 As DecoderFallbackBuffer
    Set b1 = Decoder.FallbackBuffer
    Set b2 = Decoder.FallbackBuffer
    
    Assert.That b1, Iz.Not.Nothing
    Assert.That b1, Iz.SameAs(b2)
End Sub

Public Sub Convert_WithNullBytes_ThrowsArgumentNullException()
    Dim Bytes() As Byte
    Dim Chars(0) As Integer
    Dim Decoder As Decoder
    Set Decoder = MakeEmptyDecoder

    On Error Resume Next
    Decoder.Convert Bytes, 0, 0, Chars, 0, 0, False, 0, 0, False
    
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub Convert_WithNullChars_ThrowsArgumentNullException()
    Dim Bytes(0) As Byte
    Dim Chars() As Integer
    Dim Decoder As Decoder
    Set Decoder = MakeEmptyDecoder

    On Error Resume Next
    Decoder.Convert Bytes, 0, 0, Chars, 0, 0, False, 0, 0, False
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub Convert_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Bytes(0) As Byte
    Dim Chars(0) As Integer
    Dim Decoder As Decoder
    Set Decoder = MakeEmptyDecoder

    On Error Resume Next
    Decoder.Convert Bytes, -1, 0, Chars, 0, 0, False, 0, 0, False

    AssertArgumentOutOfRangeException Err, "ByteIndex"
End Sub

Public Sub Convert_WithByteCountNegative_ThrowsArgumentOutOfRangeException()
    Dim Bytes(0) As Byte
    Dim Chars(0) As Integer
    Dim Decoder As Decoder
    Set Decoder = MakeEmptyDecoder

    On Error Resume Next
    Decoder.Convert Bytes, 0, -1, Chars, 0, 0, False, 0, 0, False

    AssertArgumentOutOfRangeException Err, "ByteCount"
End Sub

Public Sub Convert_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Bytes(0) As Byte
    Dim Chars(0) As Integer
    Dim Decoder As Decoder
    Set Decoder = MakeEmptyDecoder

    On Error Resume Next
    Decoder.Convert Bytes, 0, 0, Chars, -1, 0, False, 0, 0, False

    AssertArgumentOutOfRangeException Err, "CharIndex"
End Sub

Public Sub Convert_WithCharCountNegative_ThrowsArgumentOutOfRangeException()
    Dim Bytes(0) As Byte
    Dim Chars(0) As Integer
    Dim Decoder As Decoder
    Set Decoder = MakeEmptyDecoder

    On Error Resume Next
    Decoder.Convert Bytes, 0, 0, Chars, 0, -1, False, 0, 0, False

    AssertArgumentOutOfRangeException Err, "CharCount"
End Sub

Public Sub Convert_WithByteIndexAndByteCountMakeInvalidPosition_ThrowsArgumentOutOfRangeException()
    Dim Bytes(0) As Byte
    Dim Chars(0) As Integer
    Dim Decoder As Decoder
    Set Decoder = MakeEmptyDecoder

    On Error Resume Next
    Decoder.Convert Bytes, 0, 2, Chars, 0, 0, False, 0, 0, False

    AssertArgumentOutOfRangeException Err, "Bytes"
End Sub

Public Sub Convert_WithCharIndexAndCharCountMakeInvalidPosition_ThrowsArgumentOutOfRangeException()
    Dim Bytes(0) As Byte
    Dim Chars(0) As Integer
    Dim Decoder As Decoder
    Set Decoder = MakeEmptyDecoder

    On Error Resume Next
    Decoder.Convert Bytes, 0, 1, Chars, 0, 2, False, 0, 0, False

    AssertArgumentOutOfRangeException Err, "Chars"
End Sub

Public Sub Convert_WithSmallerChars_FillsChars()
    Dim Bytes(0 To 10000) As Byte
    Dim Chars(0 To 10000) As Integer
    Dim Decoder As Decoder
    Dim BytesUsed As Long
    Dim CharsUsed As Long
    Dim Completed As Boolean
    
    Set Decoder = Encoding.Unicode.GetDecoder
    Decoder.Convert Bytes, 0, 10000, Chars, 0, 1000, True, BytesUsed, CharsUsed, Completed
    
    Assert.That Completed, Iz.False
    Assert.That CharsUsed, Iz.EqualTo(625)
    Assert.That BytesUsed, Iz.EqualTo(1250)
End Sub


Private Function MakeEmptyDecoder() As Decoder
    Dim Enc As UnicodeEncoding
    Set Enc = Encoding.Unicode.Clone
    Set Enc.DecoderFallback = Cor.NewDecoderReplacementFallback("")
    
    Set MakeEmptyDecoder = Enc.GetDecoder
End Function











