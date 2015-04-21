VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCspParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCspParameters
'
Option Explicit

Public Sub CreateDefault()
    Dim params As New CspParameters
    Assert.AreEqual CspProviderFlags.NoFlags, params.Flags
    Assert.AreEqual "", params.KeyContainerName
    Assert.AreEqual "", params.ProviderName
    Assert.AreEqual -1, params.KeyNumber
    Assert.AreEqual 1, params.ProviderType
End Sub

Public Sub CreateWithProviderType1()
    Dim params As CspParameters
    Set params = Cor.NewCspParameters(1)
    Assert.AreEqual CspProviderFlags.NoFlags, params.Flags
    Assert.AreEqual "", params.KeyContainerName
    Assert.AreEqual "", params.ProviderName
    Assert.AreEqual -1, params.KeyNumber
    Assert.AreEqual 1, params.ProviderType
End Sub

Public Sub CreateWithProviderType2()
    Dim params As CspParameters
    Set params = Cor.NewCspParameters(2)
    Assert.AreEqual CspProviderFlags.NoFlags, params.Flags
    Assert.AreEqual "", params.KeyContainerName
    Assert.AreEqual "", params.ProviderName
    Assert.AreEqual -1, params.KeyNumber
    Assert.AreEqual 2, params.ProviderType
End Sub

Public Sub CreateWithAllParameters()
    Dim params As CspParameters
    Set params = Cor.NewCspParameters(2, "Provider", "Container")
    Assert.AreEqual CspProviderFlags.NoFlags, params.Flags
    Assert.AreEqual "Container", params.KeyContainerName
    Assert.AreEqual "Provider", params.ProviderName
    Assert.AreEqual -1, params.KeyNumber
    Assert.AreEqual 2, params.ProviderType
End Sub

