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
        Case "EncodeOID_WithValue_ReturnsExpected"
            Test.Use "1.3.14.3.2.26", NewBytes(&H6, &H5, &H2B, &HE, &H3, &H2, &H1A)
            Test.Use "2.16.840.1.101.3.4.2.1", NewBytes(&H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H1)
            Test.Use "2.16.840.1.101.3.4.2.2", NewBytes(&H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H2)
            Test.Use "2.16.840.1.101.3.4.2.3", NewBytes(&H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H3)
            Test.Use "1.3.36.3.2.1", NewBytes(&H6, &H5, &H2B, &H24, &H3, &H2, &H1)
            Test.Use "1.2.840.113549.2.5", NewBytes(&H6, &H8, &H2A, &H86, &H48, &H86, &HF7, &HD, &H2, &H5)
            
    End Select
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

Public Sub AddAlgorithm_WithEmptyTypeName_ThrowsArgumentException()
    On Error Resume Next
    
    CryptoConfig.AddAlgorithm "", "abc"
    
    AssertArgumentException Err, "TypeName"
End Sub

Public Sub AddAlgorithm_WithNonStringAlgName_ThrowsInvalidCastException()
    On Error Resume Next
    
    CryptoConfig.AddAlgorithm "Server.Class", 1234
    
    AssertInvalidCastException Err
End Sub

Public Sub AddAlgorithm_WithEmptyAlgName_ThrowsArgumentException()
    On Error Resume Next
    
    CryptoConfig.AddAlgorithm "Server.Class", ""
    
    AssertArgumentException Err
End Sub

Public Sub CreateFromName_WithCustomAlgoAdded_ReturnsCustomAlgoInstance()
    CryptoConfig.AddAlgorithm "Fakes.MockObject", "MyType"
    
    Dim Actual As Object
    Set Actual = CryptoConfig.CreateFromName("MyType")
    
    Assert.That Actual, Iz.Not.Nothing
    Assert.That TypeOf Actual Is MockObject, Iz.True
End Sub

Public Sub MapNameToOID_WithDefaultAlgorithms_ReturnsExpected()
    MapNameToOID "", "1.2.840.113549.1.9.9"
    MapNameToOID "SHA", "1.3.14.3.2.26"
    MapNameToOID "SHA1", "1.3.14.3.2.26"
    MapNameToOID "CorLib.SHA1", "1.3.14.3.2.26"
'//      MapNameToOID "CorLib.HashAlgorithm", "1.3.14.3.2.26"
    MapNameToOID "CorLib.SHA1CryptoServiceProvider", "1.3.14.3.2.26"
    MapNameToOID "CorLib.SHA1Managed", "1.3.14.3.2.26"
    MapNameToOID "MD5", "1.2.840.113549.2.5"
    MapNameToOID "CorLib.MD5", "1.2.840.113549.2.5"
    MapNameToOID "CorLib.MD5CryptoServiceProvider", "1.2.840.113549.2.5"
    MapNameToOID "SHA256", "2.16.840.1.101.3.4.2.1"
    MapNameToOID "CorLib.SHA256", "2.16.840.1.101.3.4.2.1"
    MapNameToOID "CorLib.SHA256Managed", "2.16.840.1.101.3.4.2.1"
    MapNameToOID "SHA384", "2.16.840.1.101.3.4.2.2"
    MapNameToOID "CorLib.SHA384", "2.16.840.1.101.3.4.2.2"
    MapNameToOID "CorLib.SHA384Managed", "2.16.840.1.101.3.4.2.2"
    MapNameToOID "SHA512", "2.16.840.1.101.3.4.2.3"
    MapNameToOID "CorLib.SHA512", "2.16.840.1.101.3.4.2.3"
    MapNameToOID "CorLib.SHA512Managed", "2.16.840.1.101.3.4.2.3"
    MapNameToOID "RIPEMD160", "1.3.36.3.2.1"
    MapNameToOID "CorLib.RIPEMD160", "1.3.36.3.2.1"
    MapNameToOID "CorLib.RIPEMD160Managed", "1.3.36.3.2.1"
'        // LAMESPEC: only documentated in ".NET Framework Security" book
    MapNameToOID "TripleDESKeyWrap", "1.2.840.113549.1.9.16.3.6"
'        // new OID defined in Fx 2.0
    MapNameToOID "DES", "1.3.14.3.2.7"
    MapNameToOID "TripleDES", "1.2.840.113549.3.7"
    MapNameToOID "RC2", "1.2.840.113549.3.2"
    MapNameToOID "RSA", ""
    MapNameToOID "DSA", ""
    MapNameToOID "3DES", ""
'        // no OID defined ?
    MapNameToOID "CorLib.RSA", ""
    MapNameToOID "CorLib.AsymmetricAlgorithm", ""
    MapNameToOID "CorLib.DSA", ""
    MapNameToOID "CorLib.DES", ""
    MapNameToOID "Triple DES", ""
    MapNameToOID "CorLib.TripleDES", ""
    MapNameToOID "CorLib.RC2", ""
    MapNameToOID "Rijndael", ""
    MapNameToOID "CorLib.Rijndael", ""
    MapNameToOID "CorLib.SymmetricAlgorithm", ""
'        // LAMESPEC Undocumented Names in CryptoConfig
    MapNameToOID "RandomNumberGenerator", ""
    MapNameToOID "CorLib.RandomNumberGenerator", ""
    MapNameToOID "CorLib.KeyedHashAlgorithm", ""
    MapNameToOID "HMAC", ""
    MapNameToOID "CorLib.HMAC", ""
    MapNameToOID "HMACMD5", ""
    MapNameToOID "CorLib.HMACMD5", ""
    MapNameToOID "HMACRIPEMD160", ""
    MapNameToOID "CorLib.HMACRIPEMD160", ""
    MapNameToOID "HMACSHA256", ""
    MapNameToOID "CorLib.HMACSHA256", ""
    MapNameToOID "HMACSHA384", ""
    MapNameToOID "CorLib.HMACSHA384", ""
    MapNameToOID "HMACSHA512", ""
    MapNameToOID "CorLib.HMACSHA512", ""
    MapNameToOID "HMACSHA1", ""
    MapNameToOID "CorLib.HMACSHA1", ""
    MapNameToOID "MACTripleDES", ""
    MapNameToOID "CorLib.MACTripleDES", ""
'        // non existing algo should return null (without exception)
    MapNameToOID "NonExistingAlgorithm", ""
End Sub


Private Sub CreateFromName(ByRef Name As String, ByRef ObjectName As String)
    Dim Actual As Object
    Set Actual = CryptoConfig.CreateFromName(Name)
    
    If Len(ObjectName) = 0 Then
        Assert.That Actual, Iz.Nothing
    Else
        Assert.That Actual, Iz.Not.Nothing, Name
        Assert.That Actual.ToString, Iz.EqualTo(ObjectName), Name
    End If
End Sub

Private Sub MapNameToOID(ByRef Name As String, ByRef Oid As String)
    Assert.That CryptoConfig.MapNameToOID(Name), Iz.EqualTo(Oid), "oid(" & Name & ")"
End Sub
