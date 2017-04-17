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
            
        Case "T3_Malformed_3_LastContinuationMissing_Pass"
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
            
    End Select
End Sub


Public Sub Setup()
    Set mUTF8 = NewUTF8Encoding(True, True)
End Sub

Public Sub GetByteCount_WithFailingFallback_ReturnsExpected(ByRef s As String, ByVal Expected As Long)
    Dim Enc As UTF8Encoding
    Set Enc = Encoding.UTF8.Clone
    Set Enc.EncoderFallback = New NullEncoderFallback
    
    Dim Actual As Long
    Actual = Enc.GetByteCount(s)
    
    Assert.That Actual, Iz.EqualTo(Expected)
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

Public Sub GetBytes_WithTwoHighSurrogates_ThrowsArgumentOutOfRangeException()
    Dim Encoding As UTF8Encoding
    Set Encoding = NewUTF8Encoding(False, True)
    
    On Error Resume Next
    Encoding.GetBytes t("\uD800\uD800")
    
    AssertEncoderFallbackException Err
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

Public Sub T1_Correct_GreekWord_kosme()
    Dim Data() As Byte
    Dim s As String
    
    Data = NewBytes(&HCE, &HBA, &HE1, &HBD, &HB9, &HCF, &H83, &HCE, &HBC, &HCE, &HB5)
    s = mUTF8.GetString(Data)
    
    Assert.That BitConverter.ToString(mUTF8.GetBytes(s)), Iz.EqualTo(BitConverter.ToString(Data))
End Sub

Public Sub T2_Boundary_1_FirstPossibleSequence_Pass(ByRef Data() As Byte, ByVal Expected As Long)
    Dim s As String
    
    s = mUTF8.GetString(Data)
    
    Assert.That MemLong(StrPtr(s)), Iz.EqualTo(Expected)
    Assert.That mUTF8.GetBytes(s), Iz.EqualTo(Data)
End Sub

Public Sub T2_Boundary_1_FirstPossibleSequence_Fail_5()
    Dim data215() As Byte
    data215 = NewBytes(&HF8, &H88, &H80, &H80, &H80)
    
    On Error Resume Next
    Dim Actual As String
    Actual = mUTF8.GetString(data215)
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T2_Boundary_2_FirstPossibleSequence__Fail_6()
    Dim Data() As Byte
    Data = NewBytes(&HFC, &H84, &H80, &H80, &H80, &H80)
    
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T2_Boundary_2_LastPossibleSequence_Pass(ByRef Data() As Byte, ByVal Expected As Long)
    Dim s As String
    
    s = mUTF8.GetString(Data)
    
    Assert.That MemLong(StrPtr(s)), Iz.EqualTo(Expected)
    Assert.That mUTF8.GetBytes(s), Iz.EqualTo(Data)
End Sub

Public Sub T2_Boundary_2_LastPossibleSequence_Fail_4()
    Dim Data() As Byte
    Data = NewBytes(&H7F, &HBF, &HBF, &HBF)
    
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T2_Boundary_2_LastPossibleSequence_Fail_5()
    Dim Data() As Byte
    Data = NewBytes(&HFB, &HBF, &HBF, &HBF, &HBF)
    
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T2_Boundary_2_LastPossibleSequence_Fail_6()
    Dim Data() As Byte
    Data = NewBytes(&HFB, &HBF, &HBF, &HBF, &HBF, &HBF)
    
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T2_Boundary_3_Other_Pass(ByRef Data() As Byte, ByVal Expected As Long)
    Dim s As String
    
    s = mUTF8.GetString(Data)
    
    Assert.That MemLong(StrPtr(s)), Iz.EqualTo(Expected)
    Assert.That mUTF8.GetBytes(s), Iz.EqualTo(Data)
End Sub

Public Sub T2_Boundary_3_Other_Fail_5()
    Dim Data() As Byte
    Data = NewBytes(&HF4, &H90, &H80, &H80)
    
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T3_Malformed_1_UnexpectedContinuation_3(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T3_Malformed_2_LonelyStart_3(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T3_Malformed_3_LastContinuationMissing(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T3_Malformed_3_LastContinuationMissing_Pass(ByRef Data() As Byte)
    Dim Actual As String
     
    Actual = mUTF8.GetString(Data)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub T3_Malformed_4_ConcatenationImcomplete()
    Dim Data() As Byte
    Data = NewBytes(&HC0, &HE0, &H80, &HF0, &H80, &H80, &HF8, &H80, &H80, &H80, &HFC, &H80, &H80, &H80, &H80, &HDF, _
                    &HEF, &HBF, &HF7, &HBF, &HBF, &HFB, &HBF, &HBF, &HBF, &HFD, &HBF, &HBF, &HBF, &HBF)
        
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T3_Malformed_5_ImpossibleBytes(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T4_Overlong_1_ASCII_Slash(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T4_Overlong_2_MaximumBoundary(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T4_Overlong_3_NUL(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T5_IllegalCodePosition_1_UTF16Surrogates(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub T5_IllegalCodePosition_2_PairedUTF16Surrogates(ByRef Data() As Byte)
    On Error Resume Next
    mUTF8.GetString Data
    
    AssertArgumentException Err, "Bytes"
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
    
    AssertArgumentException Err, "Bytes"
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









