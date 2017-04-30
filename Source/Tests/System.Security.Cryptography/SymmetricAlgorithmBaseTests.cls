VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SymmetricAlgorithmBaseTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' SymmetricAlgorithmBaseTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub AddLegalBlockSize_WhenCalled_AddsSizes()
    Dim Base As New SymmetricAlgorithmBase
    
    Base.AddLegalBlockSize 64, 64, 0
    Base.AddLegalBlockSize 128, 192, 64
    
    Dim b() As KeySizes
    b = Base.LegalBlockSizes
    Assert.That CorArray.Length(b), Iz.EqualTo(2)
    Assert.That b(0).MinSize, Iz.EqualTo(64)
    Assert.That b(0).MaxSize, Iz.EqualTo(64)
    Assert.That b(0).SkipSize, Iz.EqualTo(0)
    Assert.That b(1).MinSize, Iz.EqualTo(128)
    Assert.That b(1).MaxSize, Iz.EqualTo(192)
    Assert.That b(1).SkipSize, Iz.EqualTo(64)
End Sub

Public Sub BlockSize_WhenSetToLegalSize_StoresValue()
    Dim Base As New SymmetricAlgorithmBase
    Base.AddLegalBlockSize 64, 192, 64
    
    Base.BlockSize = 192
    
    Assert.That Base.BlockSize, Iz.EqualTo(192)
End Sub

Public Sub BlockSize_WhenSetToIllegalSize_ThrowsCryptographicException()
    Dim Base As New SymmetricAlgorithmBase
    Base.AddLegalBlockSize 64, 64, 0
    
    On Error Resume Next
    Base.BlockSize = 128

    AssertCryptographicException Err
End Sub

Public Sub KeySize_WhenSetToLegalSize_SetsValue()
    Dim Base As New SymmetricAlgorithmBase
    Base.AddLegalKeySize 64, 192, 64
    
    Base.KeySize = 192

    Assert.That Base.KeySize, Iz.EqualTo(192)
End Sub

Public Sub KeySize_WhenSetToIllegalValue_ThrowsCryptoGraphicException()
    Dim Base As New SymmetricAlgorithmBase
    Base.AddLegalKeySize 64, 64, 0
    
    On Error Resume Next
    Base.KeySize = 128
    
    AssertCryptographicException Err
End Sub

Public Sub GetIV_WhenUnset_GeneratesIV()
    Dim Base As New SymmetricAlgorithmBase
    Base.BlockSize = 64
    
    Dim Actual() As Byte
    Actual = Base.GetIV
    
    Assert.That CorArray.Length(Actual), Iz.EqualTo(8)
    Assert.That Actual, Iz.Not.EqualTo(NewBytes(0, 0, 0, 0, 0, 0, 0, 0))
End Sub

Public Sub GetIV_AfterGenerating_ChangesFromOriginal()
    Dim Base        As New SymmetricAlgorithmBase
    Dim Original()  As Byte
    Dim Actual()    As Byte
    
    Base.BlockSize = 64
    
    Original = Base.GetIV
    Base.GenerateIV
    Actual = Base.GetIV
    
    Assert.That Actual, Iz.Not.EqualTo(Original)
End Sub

Public Sub SetIV_WhenCalled_StoresValue()
    Dim Base As New SymmetricAlgorithmBase
    Dim IV() As Byte
    IV = NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    Base.BlockSize = 64
    
    Base.SetIV IV
    
    Assert.AreEqual IV, Base.GetIV
End Sub

Public Sub IV_WhenSetToInvalidLength_ThrowsCryptographicException()
    Dim Base As New SymmetricAlgorithmBase
    Base.BlockSize = 64
    
    On Error Resume Next
    Base.SetIV NewBytes(1, 2, 3, 4)

    AssertCryptographicException Err
End Sub

Public Sub IV_WhenSettingBlockSize_ShouldChange()
    Dim Base As New SymmetricAlgorithmBase
    Dim IV() As Byte
    
    Base.BlockSize = 64
    IV = Base.GetIV
    Base.BlockSize = 64
    
    Assert.That IV, Iz.Not.EqualTo(Base.GetIV)
End Sub

Public Sub Mode_WhenSetToInvalidCipherMode_ThrowsArgumentOutOfRangeException()
    Dim Base As New SymmetricAlgorithmBase
    
    On Error Resume Next
    Base.Mode = CTS
    
    AssertCryptographicException Err
End Sub
