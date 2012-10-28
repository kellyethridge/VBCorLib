VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDESCryptoServiceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDESCryptoServiceProvider
'
Option Explicit

Private csp As DESCryptoServiceProvider

Public Sub Setup()
    Set csp = New DESCryptoServiceProvider
End Sub

Public Sub AutoGenerateKeyAndIV()
    Dim b() As Byte
    b = csp.Key
    Assert.IsFalse cArray.IsNull(b), "Key should not be null."
    Assert.AreEqual 8, cArray.GetLength(b), "Key should be 8 bytes."
    
    Erase b
    b = csp.IV
    Assert.IsFalse cArray.IsNull(b), "IV should not be null."
    Assert.AreEqual 8, cArray.GetLength(b), "IV should be 8 bytes."
End Sub

Public Sub SetKeyToNull()
    Dim b() As Byte
    
    On Error GoTo errTrap
    csp.Key = b
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentNullException, "Wrong exception type caught."
End Sub

Public Sub SetKeyToInvalidSize()
    Dim b() As Byte
    ReDim b(0)
    
    On Error GoTo errTrap
    csp.Key = b
    On Error GoTo 0
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub SetKey()
    Dim Key() As Byte
    Key = NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    csp.Key = Key
    Assert.AreEqual Key, csp.Key
End Sub

Public Sub SetIV()
    Dim IV() As Byte
    IV = NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    csp.IV = IV
    Assert.AreEqual IV, csp.IV
End Sub

Public Sub SetKeySizeToNonLegalSize()
    On Error GoTo errTrap
    csp.KeySize = 128
    On Error GoTo 0
    Assert.Fail "An exception should be raised."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is CryptographicException, "Wrong exception type caught."
End Sub

Public Sub SetIVToInvalidSize()
    Dim IV() As Byte
    ReDim IV(0 To 15)
    
    On Error GoTo errTrap
    csp.IV = IV
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is CryptographicException, "Wrong exception type caught."
End Sub
