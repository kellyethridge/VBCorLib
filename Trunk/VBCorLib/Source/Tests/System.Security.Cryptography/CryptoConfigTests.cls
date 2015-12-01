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
'
' CryptoConfigTests
'
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

' Shamelessly stolen from mono-project
Public Sub CreateFromName_WithDefaultAlgorithms_ReturnsExpected()
    CreateFromName "SHA", "CorLib.SHA1CryptoServiceProvider"
    CreateFromName "SHA1", "CorLib.SHA1CryptoServiceProvider"
    CreateFromName "CorLib.SHA1", "CorLib.SHA1CryptoServiceProvider"
    CreateFromName "CorLib.HashAlgorithm", "CorLib.SHA1CryptoServiceProvider"
    CreateFromName "CorLib.SHA1CryptoServiceProvider", "CorLib.SHA1CryptoServiceProvider"
    CreateFromName "MD5", "CorLib.MD5CryptoServiceProvider"
    CreateFromName "CorLib.MD5", "CorLib.MD5CryptoServiceProvider"
    CreateFromName "CorLib.MD5CryptoServiceProvider", "CorLib.MD5CryptoServiceProvider"
    CreateFromName "SHA256", "CorLib.SHA256Managed"
    CreateFromName "SHA-256", "CorLib.SHA256Managed"
    CreateFromName "CorLib.SHA256", "CorLib.SHA256Managed"
    CreateFromName "SHA384", "CorLib.SHA384Managed"
    CreateFromName "SHA-384", "CorLib.SHA384Managed"
    CreateFromName "CorLib.SHA384", "CorLib.SHA384Managed"
    CreateFromName "SHA512", "CorLib.SHA512Managed"
    CreateFromName "SHA-512", "CorLib.SHA512Managed"
    CreateFromName "CorLib.SHA512", "CorLib.SHA512Managed"
    CreateFromName "RSA", "CorLib.RSACryptoServiceProvider"
    CreateFromName "CorLib.RSA", "CorLib.RSACryptoServiceProvider"
    CreateFromName "CorLib.AsymmetricAlgorithm", "CorLib.RSACryptoServiceProvider"
'    CreateFromName "DSA", "CorLib.DSACryptoServiceProvider"
'    CreateFromName "CorLib.DSA", "CorLib.DSACryptoServiceProvider"
    CreateFromName "DES", "CorLib.DESCryptoServiceProvider"
    CreateFromName "CorLib.DES", "CorLib.DESCryptoServiceProvider"
    CreateFromName "3DES", "CorLib.TripleDESCryptoServiceProvider"
    CreateFromName "TripleDES", "CorLib.TripleDESCryptoServiceProvider"
    CreateFromName "Triple DES", "CorLib.TripleDESCryptoServiceProvider"
    
    CreateFromName "CorLib.SymmetricAlgorithm", "CorLib.RijndaelManaged"
    CreateFromName "RC2", "CorLib.RC2CryptoServiceProvider"
    CreateFromName "CorLib.RC2", "CorLib.RC2CryptoServiceProvider"
    CreateFromName "Rijndael", "CorLib.RijndaelManaged"
    CreateFromName "CorLib.Rijndael", "CorLib.RijndaelManaged"
    
    CreateFromName "RandomNumberGenerator", "CorLib.RNGCryptoServiceProvider"
    CreateFromName "CorLib.RandomNumberGenerator", "CorLib.RNGCryptoServiceProvider"
    CreateFromName "CorLib.KeyedHashAlgorithm", "CorLib.HMACSHA1"
    CreateFromName "HMACSHA1", "CorLib.HMACSHA1"
    CreateFromName "CorLib.HMACSHA1", "CorLib.HMACSHA1"
    CreateFromName "MACTripleDES", "CorLib.MACTripleDES"
    CreateFromName "CorLib.MACTripleDES", "CorLib.MACTripleDES"
    
    Assert.That CryptoConfig.CreateFromName("HMAC"), Iz.Nothing, "HMAC"
    CreateFromName "CorLib.HMAC", "CorLib.HMACSHA1"
    CreateFromName "HMACMD5", "CorLib.HMACMD5"
    CreateFromName "CorLib.HMACMD5", "CorLib.HMACMD5"
    CreateFromName "HMACRIPEMD160", "CorLib.HMACRIPEMD160"
    CreateFromName "CorLib.HMACRIPEMD160", "CorLib.HMACRIPEMD160"
    CreateFromName "HMACSHA256", "CorLib.HMACSHA256"
    CreateFromName "CorLib.HMACSHA256", "CorLib.HMACSHA256"
    CreateFromName "HMACSHA384", "CorLib.HMACSHA384"
    CreateFromName "CorLib.HMACSHA384", "CorLib.HMACSHA384"
    CreateFromName "HMACSHA512", "CorLib.HMACSHA512"
    CreateFromName "CorLib.HMACSHA512", "CorLib.HMACSHA512"
    
    CreateFromName "RIPEMD160", "CorLib.RIPEMD160Managed"
    CreateFromName "RIPEMD-160", "CorLib.RIPEMD160Managed"
    CreateFromName "CorLib.RIPEMD160", "CorLib.RIPEMD160Managed"
    
    CreateFromName "CorLib.MemoryStream", "CorLib.MemoryStream"
    
    Assert.That CryptoConfig.CreateFromName("NonExistingAlgorithm"), Iz.Nothing
End Sub

Public Sub CreateFromName_WithURL_ReturnsExpected()
    CreateFromName "http://www.w3.org/2000/09/xmldsig#hmac-sha1", ""
    ' URL used in DigestMethod element
    CreateFromName "http://www.w3.org/2000/09/xmldsig#sha1", "CorLib.SHA1CryptoServiceProvider"
    ' URL used in Reference element
    CreateFromName "http://www.w3.org/2000/09/xmldsig#Object", ""
    CreateFromName "http://www.w3.org/2000/09/xmldsig#Manifest", ""
    CreateFromName "http://www.w3.org/2000/09/xmldsig#SignatureProperties", ""
    ' LAMESPEC: only documentated in ".NET Framework Security" book
    CreateFromName "http://www.w3.org/2001/04/xmlenc#sha256", "CorLib.SHA256Managed"
    CreateFromName "http://www.w3.org/2001/04/xmlenc#sha384", ""
    CreateFromName "http://www.w3.org/2001/04/xmlenc#sha512", "CorLib.SHA512Managed"

    CreateFromName "http://www.w3.org/2001/04/xmldsig-more#hmac-sha256", "CorLib.HMACSHA256"
    CreateFromName "http://www.w3.org/2001/04/xmldsig-more#hmac-sha384", "CorLib.HMACSHA384"
    CreateFromName "http://www.w3.org/2001/04/xmldsig-more#hmac-sha512", "CorLib.HMACSHA512"
    CreateFromName "http://www.w3.org/2001/04/xmldsig-more#hmac-ripemd160", "CorLib.HMACRIPEMD160"
End Sub


Private Sub CreateFromName(ByRef Name As String, ByRef ObjectName As String)
    Dim Actual As IObject
    Set Actual = CryptoConfig.CreateFromName(Name)
    
    If Len(ObjectName) = 0 Then
        Assert.That Actual, Iz.Nothing
    Else
        Assert.That Actual, Iz.Not.Nothing, Name
        Assert.That Actual.ToString, Iz.EqualTo(ObjectName), Name
    End If
End Sub

