VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CspParametersTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CspParametersTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Constructor_WhenCalled_SetsDefaultValues()
    Dim Params As New CspParameters
    
    Assert.That Params.ProviderType, Iz.EqualTo(PROV_RSA_AES)
    Assert.That Params.KeyContainerName, Iz.Empty
    Assert.That Params.ProviderName, Iz.Empty
    Assert.That Params.KeyNumber, Iz.EqualTo(-1)
End Sub

Public Sub Constructor_WithValues_SetsProperties()
    Dim Params As CspParameters
    
    Set Params = Cor.NewCspParameters(5, "Provider", "Container")
    
    Assert.That Params.ProviderType, Iz.EqualTo(5)
    Assert.That Params.ProviderName, Iz.EqualTo("Provider")
    Assert.That Params.KeyContainerName, Iz.EqualTo("Container")
End Sub

