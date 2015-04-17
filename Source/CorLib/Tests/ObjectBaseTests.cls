VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectBaseTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ObjectBase Tests
'
Option Explicit

Public Sub Equals_WithNothingAndNothing_ReturnsTrue()
    Dim Actual As Boolean
    
    Actual = MyBase.Equals(Nothing, Nothing)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Equals_WithNothingAndNotNothing_ReturnsFalse()
    Dim Actual As Boolean
    
    Actual = MyBase.Equals(Nothing, "a")
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equals_WithSameInstance_ReturnsTrue()
    Dim Obj As New Int32
    Dim Actual As Boolean
    
    Actual = MyBase.Equals(Obj, Obj)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Equals_WithDifferentInstances_ReturnsFalse()
    Dim ObjA As New Int32
    Dim ObjB As New Int32
    Dim Actual As Boolean
    
    Actual = MyBase.Equals(ObjA, ObjB)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub GetHashCode_WithNothing_ReturnsZero()
    Dim Actual As Long
    
    Actual = MyBase.GetHashCode(Nothing)
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub GetHashCode_WithSameInstanceTwice_ReturnsSameHashCode()
    Dim Obj As New Int32
    Dim Actual As Boolean
    
    Actual = MyBase.GetHashCode(Obj) = MyBase.GetHashCode(Obj)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub GetHashCode_WithDifferentInstances_ReturnsDifferentValues()
    Dim ObjA As New Int32
    Dim ObjB As New Int32
    Dim Actual As Boolean
    
    Actual = MyBase.GetHashCode(ObjA) = MyBase.GetHashCode(ObjB)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub ToString_WithNothing_ReturnsEmptyString()
    Dim Actual As String
    
    Actual = MyBase.ToString(Nothing)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub ToString_WithObjectInstance_ReturnsObjectTypeName()
    Dim Actual As String
    Dim Obj As New Int32
    
    Actual = MyBase.ToString(Obj)
    
    Assert.That Actual, Iz.EqualTo("Int32")
End Sub

Public Sub ToString_WithObjectAndApp_ReturnsFullTypeName()
    Dim Actual As String
    Dim Obj As New Int32
    
    Actual = MyBase.ToString(Obj, App)
    
    Assert.That Actual, Iz.EqualTo("VBCorLibTests.Int32")
End Sub

Public Sub ToString_WithNothingTypeAndApp_ReturnsEmpty()
    Dim Actual As String
    
    Actual = MyBase.ToString(Nothing, App)
    
    Assert.That Actual, Iz.Empty
End Sub
