VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCspKeyContainerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCspKeyContainerInfo
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

'Public Sub CreateDefault()
'    Dim Params As New CspParameters
'
'    Dim info As CspKeyContainerInfo
'    Set info = Cor.NewCspKeyContainerInfo(Params)
'
'    Assert.IsTrue info.Accessible, "Should be accessible"
'    Assert.IsFalse info.Exportable, "Should not be exportable"
'    Assert.AreEqual "", info.KeyContainerName, "KeyContainerName should be empty"
'    Assert.AreEqual KeyNumber.Exchange, info.KeyNumber, "KeyNumber should be Exchange"
'    Assert.IsLike "*?_????????-????-????-????-????????????", info.UniqueKeyContainerName, "Unique name should not be empty"
'    Assert.IsFalse info.MachineKeyStore, "Machine Store should be false"
'    Assert.IsFalse info.Protected, "Protected should be false"
'    Assert.IsLike "Microsoft *?* Cryptographic Provider", info.ProviderName
'    Assert.AreEqual 1, info.ProviderType
'    Assert.IsFalse info.Removable, "Removable should be false"
'    Assert.IsFalse info.RandomlyGenerated, "Random should be false"
'End Sub

Public Sub DifferentKeyNumbers()
    Dim Params As New CspParameters
    Dim info As CspKeyContainerInfo
    
    Set info = Cor.NewCspKeyContainerInfo(Params)
    Assert.AreEqual VBCorLib.KeyNumber.Exchange, info.KeyNumber
    
    Params.KeyNumber = 2
    Set info = Cor.NewCspKeyContainerInfo(Params)
    Assert.AreEqual VBCorLib.KeyNumber.Signature, info.KeyNumber
    
    Params.KeyNumber = 3
    Set info = Cor.NewCspKeyContainerInfo(Params)
    Assert.AreEqual 3, info.KeyNumber
End Sub

'Public Sub ProviderNameLazyLoading()
'    Dim Params As New CspParameters
'    Dim info As CspKeyContainerInfo
'    Dim s As String
'
'
'    Set info = Cor.NewCspKeyContainerInfo(Params)
'    Assert.AreEqual "", info.ProviderName
'    s = info.ProviderName
'    Assert.AreEqual "", info.ProviderName
'
'    Set info = Cor.NewCspKeyContainerInfo(Params)
'    Assert.AreEqual "", info.ProviderName
'    s = info.Accessible
'    Assert.AreNotEqual "", info.ProviderName
'
'    Set info = Cor.NewCspKeyContainerInfo(Params)
'    Assert.AreEqual "", info.ProviderName
'    s = info.Exportable
'    Assert.AreNotEqual "", info.ProviderName
'
'    Set info = Cor.NewCspKeyContainerInfo(Params)
'    Assert.AreEqual "", info.ProviderName
'    s = info.UniqueKeyContainerName
'    Assert.AreNotEqual "", info.ProviderName
'
'    Set info = Cor.NewCspKeyContainerInfo(Params)
'    Assert.AreEqual "", info.ProviderName
'    s = info.Protected
'    Assert.AreNotEqual "", info.ProviderName
'End Sub
