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
' TestUnicodeEncoding
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
        
        Case "GetByteCount_WithValues_ReturnsExpected"
            Test.Use t("za"), 4
            Test.Use t("\xd800\xdc00"), 4
            Test.Use t("\xd800a"), 4
            Test.Use t("\xd800\xd800"), 4
            Test.Use t("\xd800\xd800a"), 6
            Test.Use t("A\xdc00"), 4
            
        Case "GetByteCount_WithValues_ReturnsExpected2"
            Test.Use t("za"), 4
            Test.Use t("\xd800\xdc00"), 4
            Test.Use t("\xd800a"), 2
            Test.Use t("\xd800\xd800"), 0
            Test.Use t("\xd800\xd800a"), 2
        
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

Public Sub GetBytes_WithLittleEndianEncoding_ReturnsExpected(ByRef Chars As Variant, ByRef Expected() As Byte)
    Dim Enc As New UnicodeEncoding
    
    Dim Actual() As Byte
    Actual = Enc.GetBytes(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithBigEndianEncoding_ReturnsEncoded(ByRef Chars As Variant, ByRef Expected() As Byte)
    Dim Enc As UnicodeEncoding
    Set Enc = NewUnicodeEncoding(True, True)
    
    Dim Actual() As Byte
    Actual = Enc.GetBytes(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithValues_ReturnsExpected(ByRef Bytes() As Byte, ByRef Expected() As Integer)
    Dim Enc As New UnicodeEncoding
    Dim Actual() As Integer
    
    Actual = Enc.GetChars(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetCharsEx_WithSmallOutputBuffer_ThrowsArgumentException()
    Dim Enc As New UnicodeEncoding
    Dim SmallOutput(0) As Integer
    Dim Bytes(0 To 3) As Byte
    
    On Error Resume Next
    Enc.GetCharsEx Bytes, 0, 4, SmallOutput, 0
    
    AssertArgumentException Err, "Chars"
End Sub

Public Sub GetCharCount_WhenCalled_ReturnsExpectedCount()
    Dim b() As Byte: b = NewBytes(255, 254, 115, 0, 104, 0, 105, 0)
    Dim Enc As New UnicodeEncoding
    
    Dim Actual As Long
    Actual = Enc.GetCharCount(b, 2, CorArray.Length(b) - 2)
    
    Assert.That Actual, Iz.EqualTo(3)
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

Public Sub GetCharCount_WithValues_ReturnsExpected(ByRef Bytes() As Byte, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = Encoding.Unicode.GetCharCount(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithValues_ReturnsExpected(ByRef Chars As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = MakeEncoding.GetByteCount(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetByteCount_WithValues_ReturnsExpected2(ByRef Chars As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = MakeEncoding(Replacement:="").GetByteCount(Chars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


Private Function MakeEncoding(Optional ByVal BigEndian As Boolean, Optional ByVal Replacement As String = "?") As UnicodeEncoding
    Set MakeEncoding = Cor.NewUnicodeEncoding(BigEndian, True).Clone
    Set MakeEncoding.EncoderFallback = Cor.NewEncoderReplacementFallback(Replacement)
End Function











