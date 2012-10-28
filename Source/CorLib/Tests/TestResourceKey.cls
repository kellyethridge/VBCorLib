VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResourceKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestResourceKey
'
Option Explicit

Public Sub TestSameKey()
    Dim Key As ResourceKey
    Set Key = Cor.NewResourceKey(1, 1, 1)
    Assert.IsTrue Key.Equals(Key), "Key should equal itself."
End Sub

Public Sub TestSameName()
    Dim Key1 As ResourceKey
    Dim Key2 As ResourceKey
    
    Set Key1 = Cor.NewResourceKey("Custom", 1, 1)
    Set Key2 = Cor.NewResourceKey("Custom", 1, 1)
    Assert.IsTrue Key1.Equals(Key2), "Keys should be equal."
End Sub

Public Sub TestSameNameDifferentCase()
    Dim Key1 As ResourceKey
    Dim Key2 As ResourceKey
    
    Set Key1 = Cor.NewResourceKey("CUSTOM", 1, 1)
    Set Key2 = Cor.NewResourceKey("custom", 1, 1)
    Assert.IsTrue Key1.Equals(Key2), "Keys should be equal."
End Sub

Public Sub TestNotEqualName()
    Dim Key1 As ResourceKey
    Dim Key2 As ResourceKey
    
    Set Key1 = Cor.NewResourceKey("Name1", 1, 1)
    Set Key2 = Cor.NewResourceKey("Name2", 1, 1)
    Assert.IsFalse Key1.Equals(Key2), "Keys should not be equal."
End Sub

Public Sub TestNotEqualLanguageID()
    Dim Key1 As ResourceKey
    Dim Key2 As ResourceKey
    
    Set Key1 = Cor.NewResourceKey("Name", 1, 1)
    Set Key2 = Cor.NewResourceKey("Name", 1, 2)
    Assert.IsFalse Key1.Equals(Key2), "Keys should not be equal."
    Assert.IsFalse Key2.Equals(Key1), "Keys should not be equal."
End Sub

Public Sub TestEqualAnyLanguageID()
    Dim Key1 As ResourceKey
    Dim Key2 As ResourceKey
    
    Set Key1 = Cor.NewResourceKey("Name", 1, 1)
    Set Key2 = Cor.NewResourceKey("Name", 1, 0)
    Assert.IsTrue Key1.Equals(Key2), "Keys should be equal."
    Assert.IsTrue Key2.Equals(Key1), "Keys should be equal."
End Sub
