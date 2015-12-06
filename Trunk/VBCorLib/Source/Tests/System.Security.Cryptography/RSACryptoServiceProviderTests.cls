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
'    CspParameters csp = new CspParameters (1, null, "Mono1024");
'        // under MS a new keypair will only be generated the first time
'        rsa = new RSACryptoServiceProvider (csp);
'        // test default key size
'        Assert.AreEqual (1024, rsa.KeySize, "ConstructorCspParameters");
'        Assert.IsTrue (rsa.PersistKeyInCsp, "PersistKeyInCsp");
'        Assert.IsFalse (rsa.PublicOnly, "PublicOnly");
    Dim Params  As CspParameters
    Dim csp     As RSACryptoServiceProvider
    Set Params = Cor.NewCspParameters(1, , "CorLib1024")
    
    Set csp = Cor.NewRSACryptoServiceProvider(Parameters:=Params)
    
    Assert.That csp.KeySize, Iz.EqualTo(1024), "KeySize"
    Assert.That csp.PersistKeyInCsp, Iz.True, "PersistKeyInCsp"
    Assert.That csp.PublicOnly, Iz.False, "PublicOnly"
End Sub

