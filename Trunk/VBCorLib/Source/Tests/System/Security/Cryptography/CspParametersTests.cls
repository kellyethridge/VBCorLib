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
' TestCspParameters
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub CreateDefault()
    Dim Params As New CspParameters
    Assert.AreEqual CspProviderFlags.NoFlags, Params.Flags
    Assert.AreEqual "", Params.KeyContainerName
    Assert.AreEqual "", Params.ProviderName
    Assert.AreEqual -1, Params.KeyNumber
    Assert.AreEqual 1, Params.ProviderType
End Sub

Public Sub CreateWithProviderType1()
    Dim Params As CspParameters
    Set Params = Cor.NewCspParameters(1)
    Assert.AreEqual CspProviderFlags.NoFlags, Params.Flags
    Assert.AreEqual "", Params.KeyContainerName
    Assert.AreEqual "", Params.ProviderName
    Assert.AreEqual -1, Params.KeyNumber
    Assert.AreEqual 1, Params.ProviderType
End Sub

Public Sub CreateWithProviderType2()
    Dim Params As CspParameters
    Set Params = Cor.NewCspParameters(2)
    Assert.AreEqual CspProviderFlags.NoFlags, Params.Flags
    Assert.AreEqual "", Params.KeyContainerName
    Assert.AreEqual "", Params.ProviderName
    Assert.AreEqual -1, Params.KeyNumber
    Assert.AreEqual 2, Params.ProviderType
End Sub

Public Sub CreateWithAllParameters()
    Dim Params As CspParameters
    Set Params = Cor.NewCspParameters(2, "Provider", "Container")
    Assert.AreEqual CspProviderFlags.NoFlags, Params.Flags
    Assert.AreEqual "Container", Params.KeyContainerName
    Assert.AreEqual "Provider", Params.ProviderName
    Assert.AreEqual -1, Params.KeyNumber
    Assert.AreEqual 2, Params.ProviderType
End Sub

