VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTripleDESWeakKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTripleDESWeakKeys
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub TestWeakKey128()
    Dim b() As Byte
    b = NewArray(vbByte, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4)
    Assert.IsTrue TripleDES.IsWeakKey(b), "Key should be weak."
    
    b = NewArray(vbByte, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 5)
    Assert.IsTrue TripleDES.IsWeakKey(b), "Key should be weak."

    b = NewArray(vbByte, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 6)
    Assert.IsFalse TripleDES.IsWeakKey(b), "Key should not be weak."
End Sub

Public Sub TestWeakKey192()
    Dim b() As Byte
    b = NewArray(vbByte, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4)
    Assert.IsTrue TripleDES.IsWeakKey(b)
    
    b = NewArray(vbByte, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 5)
    Assert.IsTrue TripleDES.IsWeakKey(b)
    
    b = NewArray(vbByte, 3, 4, 3, 4, 3, 4, 3, 6, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 6)
    Assert.IsFalse TripleDES.IsWeakKey(b)
    
    b = NewArray(vbByte, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 6, 3, 4, 3, 4, 3, 4, 3, 4)
    Assert.IsFalse TripleDES.IsWeakKey(b)

    b = NewArray(vbByte, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 4, 3, 6)
    Assert.IsTrue TripleDES.IsWeakKey(b)
End Sub
