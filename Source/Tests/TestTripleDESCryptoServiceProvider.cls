VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTripleDESCryptoServiceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTripleDESCryptoServiceProvider
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub AutoGenerateKeyAndIV()
    Dim csp As New TripleDESCryptoServiceProvider
    Dim b() As Byte
    b = csp.Key
    Assert.IsFalse cArray.IsNull(b), "Key should not be null."
    Assert.AreEqual 24, cArray.GetLength(b), "Key should be 24 bytes."
    Assert.AreEqual 192, csp.KeySize
    
    Erase b
    b = csp.IV
    Assert.IsFalse cArray.IsNull(b), "IV should not be null."
    Assert.AreEqual 8, cArray.GetLength(b), "IV should be 8 bytes."
End Sub

Public Sub SetKeyToNull()
    Dim csp As New TripleDESCryptoServiceProvider
    Dim b() As Byte
    
    On Error GoTo errTrap
    csp.Key = b
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentNullException, "Wrong exception type caught."
End Sub

Public Sub SetKeyToInvalidSize()
    Dim csp As New TripleDESCryptoServiceProvider
    Dim b() As Byte
    ReDim b(0)
    
    On Error GoTo errTrap
    csp.Key = b
    On Error GoTo 0
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub SetKeyTo128Bits()
    Dim csp As New TripleDESCryptoServiceProvider
    Dim b() As Byte
    b = NewArray(vbByte, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 1, 2, 3, 4, 5, 6)
    
    csp.Key = b
    Assert.AreEqual 128, csp.KeySize, "KeySize should change to 128."
    Assert.AreEqual b, csp.Key
End Sub

Public Sub SetKeySizeTo128Bits()
    Dim csp As New TripleDESCryptoServiceProvider
    Assert.AreEqual 24, cArray.GetLength(csp.Key)
    
    csp.KeySize = 128
    Assert.AreEqual 16, cArray.GetLength(csp.Key)
End Sub
