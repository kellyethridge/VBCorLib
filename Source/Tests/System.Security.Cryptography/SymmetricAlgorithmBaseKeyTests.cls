VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SymmetricAlgorithmBaseKeyTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' SymmetricAlgorithmBaseKeyTests
'
Option Explicit
Implements ICategorizable

Private WithEvents Base As SymmetricAlgorithmBase
Attribute Base.VB_VarHelpID = -1
Private mEx As Exception
Private mWeaknessTested As Boolean

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Setup()
    ClearException
    Set Base = New SymmetricAlgorithmBase
    Base.AddLegalKeySize 64, 192, 64
    Base.KeySize = 64
    mWeaknessTested = False
    Set mEx = Nothing
End Sub

Public Sub GetKey_WhenUnset_GeneratesKey()
    Dim Key() As Byte
    
    Key = Base.GetKey
    
    Assert.That CorArray.Length(Key), Iz.EqualTo(8)
    Assert.That mWeaknessTested, Iz.True
End Sub

Public Sub GenerateKey_WhenKeyAlreadyGenerated_GeneratesNewKey()
    Dim Key1() As Byte
    Dim Key2() As Byte

    Key1 = Base.GetKey
    Base.GenerateKey
    Key2 = Base.GetKey

    Assert.That Key1, Iz.Not.EqualTo(Key2)
End Sub

Public Sub SetKey_WhenCalled_SetsKeyAndChecksWeakness()
    Dim Key() As Byte
    Key = NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Base.SetKey Key
    
    Assert.That Base.GetKey, Iz.EqualTo(Key)
    Assert.That mWeaknessTested, Iz.True
End Sub

Public Sub SetKey_WithWeakKeyFlag_ThrowsExpectedException()
    Dim Key() As Byte
    ReDim Key(7)
    Set mEx = Cor.NewCryptographicException("Weak key")

    On Error Resume Next
    Base.SetKey Key
    On Error GoTo 0
    
    Assert.That TakeException, Iz.SameAs(mEx)
    Assert.Pass
End Sub

Public Sub SetKey_WithInvalidKeyLength_ThrowsArgumentException()
    Dim Key() As Byte
    Key = NewBytes(1, 2, 3, 4, 5, 6, 7, 8, 9)

    On Error Resume Next
    Base.SetKey Key
    
    AssertArgumentException Err, "Key"
End Sub

Public Sub KeySize_WhenChanged_RegeneratesKey()
    Dim Key() As Byte
    Key = Base.GetKey
    
    Base.KeySize = 64
    
    Assert.That Key, Iz.Not.EqualTo(Base.GetKey)
End Sub

Private Sub base_CheckForWeakKey(RgbKey() As Byte, Ex As Exception)
    mWeaknessTested = True
    Set Ex = mEx
End Sub
