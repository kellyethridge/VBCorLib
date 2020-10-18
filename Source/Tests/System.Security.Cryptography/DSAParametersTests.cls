VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSAParametersTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' DSAParametersTests
'
Option Explicit

Public Sub Serialization_WithNullParameters_SerializesNullBytes()
    Dim p1 As New DSAParameters
    Dim p2 As DSAParameters
    
    Set p2 = Persist(p1)
    
    Assert.That p2.Counter, Iz.EqualTo(0)
    Assert.That CorArray.IsNull(p2.Seed), Iz.True
End Sub

Public Sub Serialization_WithXParameter_DoesNotIncludeXParameter()
    Dim p1 As New DSAParameters
    Dim p2 As DSAParameters
    
    p1.x = NewBytes(1)
    Set p2 = Persist(p1)
    
    Assert.That CorArray.IsNull(p2.x), Iz.True
End Sub
