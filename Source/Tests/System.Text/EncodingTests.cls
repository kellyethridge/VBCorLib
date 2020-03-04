VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EncodingTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetEncoding_WithCodePage_ReturnsCommonEncoding"
            Test.Use 1200, Encoding.Unicode
            Test.Use 1201, Encoding.BigEndianUnicode
            Test.Use 65000, Encoding.UTF7
            Test.Use 65001, Encoding.UTF8
            Test.Use 20127, Encoding.ASCII
            Test.Use 0, Encoding.Default
            
        Case "GetEncoding_WithEncodingName_ReturnsCommonEncoding"
            Test.Use "unicode", Encoding.Unicode
            Test.Use "utf-8", Encoding.UTF8
            Test.Use "us-ascii", Encoding.ASCII
            Test.Use "utf-7", Encoding.UTF7
            
    End Select
End Sub

Public Sub Unicode_WhenCalled_ReturnsLittleEndianUnicodeEncoding()
    Dim Actual As Encoding
    
    Set Actual = Encoding.Unicode
    
    Assert.That TypeOf Actual Is UnicodeEncoding, Iz.True
    Assert.That Actual.GetPreamble, Iz.EqualTo(NewBytes(&HFF, &HFE))
End Sub

Public Sub BigEndianUnicode_WhenCalled_ReturnsBigEndianUnicodeEncoding()
    Dim Actual As Encoding
    
    Set Actual = Encoding.BigEndianUnicode
    
    Assert.That TypeOf Actual Is UnicodeEncoding, Iz.True
    Assert.That Actual.GetPreamble, Iz.EqualTo(NewBytes(&HFE, &HFF))
End Sub

Public Sub GetEncoding_WithCodePage_ReturnsCommonEncoding(ByVal CodePage As Long, ByVal Expected As Encoding)
    Dim Actual As Encoding
    
    Set Actual = Encoding.GetEncoding(CodePage)
    
    Assert.That Actual, Iz.SameAs(Expected)
End Sub

Public Sub GetEncoding_WithEncodingName_ReturnsCommonEncoding(ByVal Name As String, ByVal Expected As Encoding)
    Dim Actual As Encoding
    
    Set Actual = Encoding.GetEncoding(Name)
    
    Assert.That Actual, Iz.SameAs(Expected)
End Sub
