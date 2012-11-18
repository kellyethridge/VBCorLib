VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSymmetricAlgorithmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSymmetricAlgorithmBase
'
Option Explicit

Public Sub AddLegalBlockSizes()
    Dim base As New SymmetricAlgorithmBase
    base.AddLegalBlockSize 64, 64, 0
    base.AddLegalBlockSize 128, 192, 64
    
    Dim b() As KeySizes
    b = base.LegalBlockSizes
    Assert.AreEqual 2, cArray.GetLength(b)
    Assert.AreEqual 64, b(0).MinSize
    Assert.AreEqual 64, b(0).MaxSize
    Assert.AreEqual 0, b(0).SkipSize
    Assert.AreEqual 128, b(1).MinSize
    Assert.AreEqual 192, b(1).MaxSize
    Assert.AreEqual 64, b(1).SkipSize
End Sub

Public Sub SetValidBlockSize()
    Dim base As New SymmetricAlgorithmBase
    base.AddLegalBlockSize 64, 192, 64
    
    On Error GoTo errTrap
    base.BlockSize = 192
    Assert.AreEqual 192, base.BlockSize
    Exit Sub
    
errTrap:
    Assert.Fail "BlockSize should be legal."
End Sub

Public Sub SetInvalidBlockSize()
    Dim base As New SymmetricAlgorithmBase
    base.AddLegalBlockSize 64, 64, 0
    
    On Error GoTo errTrap
    base.BlockSize = 128
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is CryptographicException, "Wrong exception type."
End Sub

Public Sub AddLegalKeySizes()
    Dim base As New SymmetricAlgorithmBase
    base.AddLegalKeySize 64, 64, 0
    base.AddLegalKeySize 128, 192, 64
    
    Dim b() As KeySizes
    b = base.LegalKeySizes
    Assert.AreEqual 2, cArray.GetLength(b)
    Assert.AreEqual 64, b(0).MinSize
    Assert.AreEqual 64, b(0).MaxSize
    Assert.AreEqual 0, b(0).SkipSize
    Assert.AreEqual 128, b(1).MinSize
    Assert.AreEqual 192, b(1).MaxSize
    Assert.AreEqual 64, b(1).SkipSize
End Sub

Public Sub SetValidKeySize()
    Dim base As New SymmetricAlgorithmBase
    base.AddLegalKeySize 64, 192, 64
    
    On Error GoTo errTrap
    base.KeySize = 192
    Assert.AreEqual 192, base.KeySize
    Exit Sub
    
errTrap:
    Assert.Fail "KeySize should be legal."
End Sub

Public Sub SetInvalidKeySize()
    Dim base As New SymmetricAlgorithmBase
    base.AddLegalKeySize 64, 64, 0
    
    On Error GoTo errTrap
    base.KeySize = 128
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is CryptographicException, "Wrong exception type."
End Sub

Public Sub GetIVWithOutSettingFirst()
    Dim base As New SymmetricAlgorithmBase
    base.BlockSize = 64
    
    Dim IV() As Byte
    IV = base.GetIV
    Assert.AreEqual 8, cArray.GetLength(IV)
End Sub

Public Sub GetIVBeforeAndAfterGenerating()
    Dim base As New SymmetricAlgorithmBase
    Dim IV1() As Byte
    Dim IV2() As Byte
    
    base.BlockSize = 64
    
    IV1 = base.GetIV
    base.GenerateIV
    IV2 = base.GetIV
    
    Assert.That IV1, Iz.Not.EqualTo(IV2)
End Sub

Public Sub SetIV()
    Dim base As New SymmetricAlgorithmBase
    Dim IV() As Byte
    IV = NewArray(vbByte, 1, 2, 3, 4, 5, 6, 7, 8)
    base.BlockSize = 64
    base.SetIV IV
    
    Assert.AreEqual IV, base.GetIV
End Sub

Public Sub SetInvalidIVLength()
    Dim base As New SymmetricAlgorithmBase
    base.BlockSize = 64
    
    Dim IV() As Byte
    IV = NewArray(vbByte, 1, 2, 3, 4)
    
    On Error GoTo errTrap
    base.SetIV IV
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is CryptographicException, "Wrong exception type."
End Sub

Public Sub SettingBlockSizeShouldChangeIV()
    Dim base As New SymmetricAlgorithmBase
    base.BlockSize = 64
    
    Dim IV() As Byte
    IV = base.GetIV
    base.BlockSize = 64
    Assert.That IV, Iz.Not.EqualTo(base.GetIV)
End Sub

Public Sub SetInvalidCipherMode()
    Dim base As New SymmetricAlgorithmBase
    On Error GoTo errTrap
    base.Mode = CTS
    On Error GoTo 0
    Assert.Fail "An exception should be raised."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is CryptographicException, "Wrong exception type caught."
End Sub
