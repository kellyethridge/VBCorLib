VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoConfigTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "MapNameToOID_WithValue_ReturnsExpected"
            Test.Use "sha1", "1.3.14.3.2.26"
            Test.Use "sha", "1.3.14.3.2.26"
            Test.Use "sha256", "2.16.840.1.101.3.4.2.1"
            Test.Use "sha384", "2.16.840.1.101.3.4.2.2"
            Test.Use "sha512", "2.16.840.1.101.3.4.2.3"
            Test.Use "md5", "1.2.840.113549.2.5"
            Test.Use "ripemd160", "1.3.36.3.2.1"
            Test.Use "", "1.2.840.113549.1.9.9"
        
        Case "EncodeOID_WithValue_ReturnsExpected"
            Test.Use "1.3.14.3.2.26", NewBytes(&H6, &H5, &H2B, &HE, &H3, &H2, &H1A)
            Test.Use "2.16.840.1.101.3.4.2.1", NewBytes(&H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H1)
            Test.Use "2.16.840.1.101.3.4.2.2", NewBytes(&H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H2)
            Test.Use "2.16.840.1.101.3.4.2.3", NewBytes(&H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H3)
            Test.Use "1.3.36.3.2.1", NewBytes(&H6, &H5, &H2B, &H24, &H3, &H2, &H1)
            Test.Use "1.2.840.113549.2.5", NewBytes(&H6, &H8, &H2A, &H86, &H48, &H86, &HF7, &HD, &H2, &H5)
            
    End Select
End Sub

Public Sub MapNameToOID_WithValue_ReturnsExpected(ByVal Name As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CryptoConfig.MapNameToOID(Name)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub EncodeOID_WithValue_ReturnsExpected(ByVal Oid As String, ByRef Expected As Variant)
    Dim Actual() As Byte
    
    Actual = CryptoConfig.EncodeOID(Oid)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


