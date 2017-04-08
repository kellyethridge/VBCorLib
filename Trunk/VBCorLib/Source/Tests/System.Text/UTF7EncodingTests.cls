VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTF7EncodingTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' UTF7EncodingTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetByteCount_WithValues_ReturnsExpected"
            Test.Use t("A\x2262\x0391."), 10
            Test.Use t("Hi Mom -\x263A-!"), 19
            Test.Use t("\x65E5\x672C\x8A9E"), 10
            Test.Use t("Hi Mom \x263A!"), 15
            
    End Select
End Sub

Public Sub GetByteCount_WithValues_ReturnsExpected(ByRef s As String, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = Encoding.UTF7.GetByteCount(s)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsBrowswerDisplay_WhenCalled_ReturnsFalse()
    Assert.That Encoding.UTF7.IsBrowserDisplay, Iz.False
End Sub

Public Sub IsBrowserSave_WhenCalled_ReturnsFalse()
    Assert.That Encoding.UTF7.IsBrowserSave, Iz.False
End Sub

Public Sub IsMailNewsDisplay_WhenCalled_ReturnsTrue()
    Assert.That Encoding.UTF7.IsMailNewsDisplay, Iz.True
End Sub

Public Sub IsMailNewsSave_WhenCalled_ReturnsTrue()
    Assert.That Encoding.UTF7.IsMailNewsSave, Iz.True
End Sub

Public Sub GetBytes_WithDirectlyEncodedCharacters_ReturnsExpected()
    Dim Chars       As String
    Dim Actual()    As Byte
    Dim Encoding    As New UTF7Encoding
    
    Chars = t("\u0061\u007A\u0041\u005A\u0030\u0039\u0027\u003F")
    
    Actual = Encoding.GetBytes(Chars)
    
    Dim Expected() As Byte
    Expected = NewBytes(&H61, &H7A, &H41, &H5A, &H30, &H39, &H27, &H3F)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytesEx_WithDirectlyEncodedCharacters_ReturnsExpected()
    Dim Chars()         As Integer
    Dim Actual(0 To 7)  As Byte
    Dim Count           As Long
    Dim Encoding        As New UTF7Encoding
    
    Chars = CorString.ToCharArray(t("\u0061\u007A\u0041\u005A\u0030\u0039\u0027\u003F"))
    
    Count = Encoding.GetBytesEx(Chars, 0, CorArray.Length(Chars), Actual, 0)
    
    Dim Expected() As Byte
    Expected = NewBytes(&H61, &H7A, &H41, &H5A, &H30, &H39, &H27, &H3F)
    Assert.That Count, Iz.EqualTo(8)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithOptionalEncodedCharactersAllowed_ReturnsExpected()
    Dim Encoding    As UTF7Encoding
    Dim Chars       As String
    Dim Actual()    As Byte
    
    Set Encoding = Cor.NewUTF7Encoding(True)
    Chars = t("\u0021\u0026\u002A\u003B")
    
    Actual = Encoding.GetBytes(Chars)
    
    Dim Expected() As Byte
    Expected = NewBytes(&H21, &H26, &H2A, &H3B)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithOptionalEncodedCharactersNotAllowed_ReturnsExpected()
    Dim Encoding    As UTF7Encoding
    Dim Chars       As String
    Dim Actual()    As Byte
    
    Set Encoding = Cor.NewUTF7Encoding(False)
    Chars = t("\u0021\u0026\u002A\u003B")
    
    Actual = Encoding.GetBytes(Chars)
    
    Dim Expected() As Byte
    Expected = NewBytes(&H2B, &H41, &H43, &H45, &H41, &H4A, &H67, &H41, &H71, &H41, &H44, &H73, &H2D)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithShiftedCharacters_ReturnsExpected()
    Dim Encoding    As UTF7Encoding
    Dim Chars       As String
    Dim Actual()    As Byte
    
    Set Encoding = Cor.NewUTF7Encoding(False)
    Chars = t("\u0041\u2262\u0391\u002E")
    
    Actual = Encoding.GetBytes(Chars)
    
    Dim Expected() As Byte
    Expected = NewBytes(&H41, &H2B, &H49, &H6D, &H49, &H44, &H6B, &H51, &H2D, &H2E)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytesEx_WithShiftedCharacters_ReturnsExpected()
    Dim Encoding        As UTF7Encoding
    Dim Chars()         As Integer
    Dim Actual(0 To 9)  As Byte
    Dim Count           As Long
    
    Set Encoding = Cor.NewUTF7Encoding(False)
    Chars = CorString.ToCharArray(t("\u0041\u2262\u0391\u002E"))
    
    Count = Encoding.GetBytesEx(Chars, 0, CorArray.Length(Chars), Actual, 0)
    
    Dim Expected() As Byte
    Expected = NewBytes(&H41, &H2B, &H49, &H6D, &H49, &H44, &H6B, &H51, &H2D, &H2E)
    Assert.That Count, Iz.EqualTo(10)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithRFC1642Example_ReturnsExpected()
    Dim Bytes()     As Byte
    Dim Actual()    As Integer
    Dim Encoding    As New UTF7Encoding
    
    Bytes = NewBytes(&H41, &H2B, &H49, &H6D, &H49, &H44, &H6B, &H51, &H2D, &H2E)
    
    Actual = Encoding.GetChars(Bytes)
    
    Dim Expected() As Integer
    Expected = NewIntegers(&H41, &H2262, &H391, &H2E)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetString_WithRFC1642Example_ReturnsExpected()
    Dim Bytes()     As Byte
    Dim Actual      As String
    Dim Encoding    As New UTF7Encoding
    
    Bytes = NewBytes(&H41, &H2B, &H49, &H6D, &H49, &H44, &H6B, &H51, &H2D, &H2E)
    
    Actual = Encoding.GetString(Bytes)
    
    Dim Expected As String
    Expected = t("\u0041\u2262\u0391\u002E")
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithRFC1642Example_ReturnsExpected2()
    Dim Bytes()     As Byte
    Dim Actual()    As Integer
    Dim Encoding    As New UTF7Encoding
    
    Bytes = NewBytes(&H48, &H69, &H20, &H4D, &H6F, &H4D, &H20, &H2B, &H4A, &H6A, &H6F, &H41, &H49, &H51, &H2D)
    
    Actual = Encoding.GetChars(Bytes)
    
    Dim Expected() As Integer
    Expected = NewIntegers(&H48, &H69, &H20, &H4D, &H6F, &H4D, &H20, &H263A, &H21)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetString_WithRFC1642Example_ReturnsExpected2()
    Dim Bytes()     As Byte
    Dim Actual      As String
    Dim Encoding    As New UTF7Encoding
    
    Bytes = NewBytes(&H48, &H69, &H20, &H4D, &H6F, &H4D, &H20, &H2B, &H4A, &H6A, &H6F, &H41, &H49, &H51, &H2D)
    
    Actual = Encoding.GetString(Bytes)
    
    Dim Expected As String
    Expected = t("\u0048\u0069\u0020\u004D\u006F\u004D\u0020\u263A\u0021")
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithRFC1642Example_ReturnsExpected3()
    Dim Bytes()     As Byte
    Dim Actual()    As Integer
    Dim Encoding    As New UTF7Encoding
    
    Bytes = NewBytes(&H2B, &H5A, &H65, &H56, &H6E, &H4C, &H49, &H71, &H65, &H2D)
    
    Actual = Encoding.GetChars(Bytes)
    
    Dim Expected() As Integer
    Expected = NewIntegers(&H65E5, &H672C, &H8A9E)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetString_WithRFC1642Example_ReturnsExpected3()
    Dim Bytes()     As Byte
    Dim Actual      As String
    Dim Encoding    As New UTF7Encoding
    
    Bytes = NewBytes(&H2B, &H5A, &H65, &H56, &H6E, &H4C, &H49, &H71, &H65, &H2D)
    
    Actual = Encoding.GetString(Bytes)
    
    Dim Expected As String
    Expected = t("\u65E5\u672C\u8A9E")
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetChars_WithRFC1642Example_ReturnsExpected4()
    Dim Bytes()     As Byte
    Dim Actual()    As Integer
    Dim Encoding    As New UTF7Encoding
    
    Bytes = NewBytes(&H49, &H74, &H65, &H6D, &H20, &H33, &H20, &H69, &H73, &H20, &H2B, &H41, &H4B, &H4D, &H2D, &H31, &H2E)
    
    Actual = Encoding.GetChars(Bytes)
    
    Dim Expected() As Integer
    Expected = NewIntegers(&H49, &H74, &H65, &H6D, &H20, &H33, &H20, &H69, &H73, &H20, &HA3, &H31, &H2E)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetString_WithRFC1642Example_ReturnsExpected4()
    Dim Bytes()     As Byte
    Dim Actual      As String
    Dim Encoding    As New UTF7Encoding
    
    Bytes = NewBytes(&H49, &H74, &H65, &H6D, &H20, &H33, &H20, &H69, &H73, &H20, &H2B, &H41, &H4B, &H4D, &H2D, &H31, &H2E)
    
    Actual = Encoding.GetString(Bytes)
    
    Dim Expected As String
    Expected = t("\u0049\u0074\u0065\u006D\u0020\u0033\u0020\u0069\u0073\u0020\u00A3\u0031\u002E")
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetMaxCharCount_With50_Returns50()
    Dim Encoding    As New UTF7Encoding
    Dim Actual      As Long
    
    Actual = Encoding.GetMaxCharCount(50)
    
    Assert.That Actual, Iz.EqualTo(50)
End Sub

Public Sub GetMaxByteCount_With50_Returns152()
    Dim Encoding    As New UTF7Encoding
    Dim Actual      As Long
    
    Actual = Encoding.GetMaxByteCount(50)
    
    Assert.That Actual, Iz.EqualTo(152)
End Sub

Public Sub GetCharCount_WithValue_ReturnsExpected()
    Const Original As String = "*123456789*123456789*123456789*123456789*123456789*123456789*123456789*123456789"
    Dim Bytes() As Byte
    
    Bytes = Encoding.UTF7.GetBytes(Original)
    Assert.That CorArray.Length(Bytes), Iz.EqualTo(112)
    Assert.That Encoding.UTF7.GetCharCount(Bytes), Iz.EqualTo(80)
    
    Dim Decoded As String
    Decoded = Encoding.UTF7.GetString(Bytes)
    Assert.That Decoded, Iz.EqualTo(Original)
End Sub























