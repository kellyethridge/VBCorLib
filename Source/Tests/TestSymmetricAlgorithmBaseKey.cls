VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSymmetricAlgorithmBaseKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSymmetricAlgorithmBaseKey
'
Option Explicit
Implements ICategorizable

Private WithEvents base As SymmetricAlgorithmBase
Attribute base.VB_VarHelpID = -1
Private mEx As Exception
Private mWeaknessTested As Boolean

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub


Public Sub Setup()
    Set base = New SymmetricAlgorithmBase
    base.AddLegalKeySize 64, 192, 64
    base.KeySize = 64
    mWeaknessTested = False
    Set mEx = Nothing
End Sub




Public Sub GetKeyWithoutSettingFirst()
    Dim Key() As Byte
    Key = base.GetKey
    Assert.AreEqual 8, cArray.GetLength(Key)
    Assert.IsTrue mWeaknessTested, "Weakness not tested."
End Sub

Public Sub GetKeyBeforeAndAfterGenerating()
    Dim Key1() As Byte
    Dim Key2() As Byte
    
    Key1 = base.GetKey
    base.GenerateKey
    Key2 = base.GetKey
    
    Assert.That Key1, Iz.Not.EqualTo(Key2)
End Sub

Public Sub SetKey()
    Dim Key() As Byte
    Key = NewArray(vbByte, 1, 2, 3, 4, 5, 6, 7, 8)
    base.SetKey Key
    Assert.AreEqual Key, base.GetKey
    Assert.IsTrue mWeaknessTested
End Sub

Public Sub SetWeakKey()
    Dim Key() As Byte
    ReDim Key(7)
    Set mEx = Cor.NewCryptographicException("Weak key")
    
    On Error GoTo errTrap
    base.SetKey Key
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Call Catch(Ex)
    Assert.AreSame mEx, Ex
End Sub

Public Sub SetInvalidKeyLength()
    Dim Key() As Byte
    Key = NewArray(vbByte, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    
    On Error GoTo errTrap
    base.SetKey Key
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub SettingKeySizeChangesKey()
    Dim Key() As Byte
    Key = base.GetKey
    base.KeySize = 64
    Assert.That Key, Iz.Not.EqualTo(base.GetKey)
End Sub



Private Sub base_CheckForWeakKey(RgbKey() As Byte, Ex As VBCorLib.Exception)
    mWeaknessTested = True
    Set Ex = mEx
End Sub
