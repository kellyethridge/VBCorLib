VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RSAParametersTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' RSAParametersTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Serialization_WhenCalled_SerializesOnlyPublicComponent()
    Dim Params As New RSAParameters
    Params.d = NewBytes(1)
    Params.DP = NewBytes(2)
    Params.DQ = NewBytes(3)
    Params.Exponent = NewBytes(4)
    Params.InverseQ = NewBytes(5)
    Params.Modulus = NewBytes(6)
    Params.p = NewBytes(7)
    Params.Q = NewBytes(8)
    
    Dim Actual As RSAParameters
    Set Actual = Persist(Params)
    
    Assert.That Actual.Exponent, Iz.EqualTo(NewBytes(4))
    Assert.That Actual.Modulus, Iz.EqualTo(NewBytes(6))
    
    ' Private component is never serialized.
    Assert.That CorArray.IsNull(Actual.d), Iz.True
    Assert.That CorArray.IsNull(Actual.DP), Iz.True
    Assert.That CorArray.IsNull(Actual.DQ), Iz.True
    Assert.That CorArray.IsNull(Actual.InverseQ), Iz.True
    Assert.That CorArray.IsNull(Actual.p), Iz.True
    Assert.That CorArray.IsNull(Actual.Q), Iz.True
End Sub
