VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RSACryptoServiceProviderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' RSACryptoServiceProviderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Constructor_WhenCalled_SetsDefaults()
    Dim csp As New RSACryptoServiceProvider
    
    Assert.That csp.KeySize, Iz.EqualTo(1024), "KeySize"
    Assert.That csp.PersistKeyInCsp, Iz.False, "PersistKeyInCsp"
    Assert.That csp.PublicOnly, Iz.False, "PublicOnly"
End Sub

Public Sub Constructor_WithKeySize_SetsKeySize()
    Dim csp As RSACryptoServiceProvider
    
    Set csp = Cor.NewRSACryptoServiceProvider(384)
    
    Assert.That csp.KeySize, Iz.EqualTo(384), "KeySize"
    Assert.That csp.PersistKeyInCsp, Iz.False, "PersistKeyInCsp"
    Assert.That csp.PublicOnly, Iz.False, "PublicOnly"
End Sub

Public Sub Constructor_WithCspParameters_SetsValues()
    Dim Params  As CspParameters
    Dim csp     As RSACryptoServiceProvider
    Set Params = Cor.NewCspParameters(1, , "CorLib1024")
    
    Set csp = Cor.NewRSACryptoServiceProvider(Parameters:=Params)
    
    Assert.That csp.KeySize, Iz.EqualTo(1024), "KeySize"
    Assert.That csp.PersistKeyInCsp, Iz.True, "PersistKeyInCsp"
    Assert.That csp.PublicOnly, Iz.False, "PublicOnly"
End Sub

Public Sub Constructor_WithKeySizeAndCspParameters_SetsValues()
    Const KeySize   As Long = 512
    Dim Params      As CspParameters
    Dim csp         As RSACryptoServiceProvider
    Set Params = Cor.NewCspParameters(1, ContainerNameIn:="CorLib512")
    
    Set csp = Cor.NewRSACryptoServiceProvider(KeySize, Params)
    
    Assert.That csp.KeySize, Iz.EqualTo(KeySize), "KeySize"
    Assert.That csp.PersistKeyInCsp, Iz.True, "PersistKeyInCsp"
    Assert.That csp.PublicOnly, Iz.False, "PublicOnly"
End Sub

Public Sub Constructor_WithTooSmallKeySize_ThrowsCryptographicException()
    Const SmallKeySize As Long = 256
    On Error Resume Next
    
    With Cor.NewRSACryptoServiceProvider(SmallKeySize)
        .ToXmlString True ' called to invoke key size validation.
    End With
        
    AssertCryptographicException Err
End Sub

Public Sub Properties_WhenCalled_ReturnsExpected()
    Dim csp As New RSACryptoServiceProvider
    Dim KeySizes() As KeySizes
    
    KeySizes = csp.LegalKeySizes
    Assert.That CorArray.Length(KeySizes), Iz.EqualTo(1), "LegalKeySizes"
    Assert.That KeySizes(0).MinSize, Iz.EqualTo(384), "LegalKeySize.MinSize"
    Assert.That KeySizes(0).MaxSize, Iz.EqualTo(16384), "LegalKeySize.MaxSize"
    Assert.That KeySizes(0).SkipSize, Iz.EqualTo(8), "LegalKeySize.SkipSize"
    Assert.That csp.KeyExchangeAlgorithm, Iz.EqualTo("RSA-PKCS1-KeyEx"), "KeyExchangeAlgorithm"
    Assert.That csp.SignatureAlgorithm, Iz.EqualTo("http://www.w3.org/2000/09/xmldsig#rsa-sha1"), "SignatureAlgorithm"
    csp.Clear
    KeySizes = csp.LegalKeySizes
    Assert.That CorArray.Length(KeySizes), Iz.EqualTo(1), "LegalKeySizes(disposed)"
    Assert.That KeySizes(0).MinSize, Iz.EqualTo(384), "LegalKeySize.MinSize(disposed)"
    Assert.That KeySizes(0).MaxSize, Iz.EqualTo(16384), "LegalKeySize.MaxSize(disposed)"
    Assert.That KeySizes(0).SkipSize, Iz.EqualTo(8), "LegalKeySize.SkipSize(disposed)"
    Assert.That csp.KeyExchangeAlgorithm, Iz.EqualTo("RSA-PKCS1-KeyEx"), "KeyExchangeAlgorithm(disposed)"
    Assert.That csp.SignatureAlgorithm, Iz.EqualTo("http://www.w3.org/2000/09/xmldsig#rsa-sha1"), "SignatureAlgorithm(disposed)"
End Sub


















