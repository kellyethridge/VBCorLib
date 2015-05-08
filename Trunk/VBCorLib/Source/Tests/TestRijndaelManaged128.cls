VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRijndaelManaged128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRijndaelManaged128
'
Option Explicit
Implements ICategorizable

Private csp As RijndaelManaged

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Setup()
    Set csp = New RijndaelManaged
    csp.KeySize = 128
End Sub


Public Sub AutoGenerateKeyAndIV()
    Dim b() As Byte
    b = csp.Key
    Assert.IsFalse CorArray.IsNull(b), "Key should not be null."
    Assert.AreEqual 16, CorArray.GetLength(b), "Key should be 16 bytes."
    
    Erase b
    b = csp.IV
    Assert.IsFalse CorArray.IsNull(b), "IV should not be null."
    Assert.AreEqual 16, CorArray.GetLength(b), "IV should be 16 bytes."
End Sub

