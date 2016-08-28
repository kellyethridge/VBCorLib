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
Option Explicit
Implements ICategorizable

Private mUTF8 As UTF8Encoding

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub Setup()
    Set mUTF8 = NewUTF8Encoding(True, True)
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













'        public void TestThrowOnInvalid ()
'        {
'            UTF8Encoding u = new UTF8Encoding (true, false);
'
'            byte[] data = new byte [] { 0xC0, 0xAF };
'            Assert.AreEqual (2, u.GetCharCount (data), "#A0");
'            string s = u.GetString (data);
'            Assert.AreEqual ("\uFFFD\uFFFD", s, "#A1");
'
'            data = new byte [] { 0x30, 0x31, 0xC0, 0xAF, 0x30, 0x32 };
'            s = u.GetString (data);
'            Assert.AreEqual (6, s.Length, "#B1");
'            Assert.AreEqual (0x30, (int) s [0], "#B2");
'            Assert.AreEqual (0x31, (int) s [1], "#B3");
'            Assert.AreEqual (0xFFFD, (int) s [2], "#B4");
'            Assert.AreEqual (0xFFFD, (int) s [3], "#B5");
'            Assert.AreEqual (0x30, (int) s [4], "#B6");
'            Assert.AreEqual (0x32, (int) s [5], "#B7");
'}























