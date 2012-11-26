VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IObject

Public Equals As Variant
Public EqualsResult As Boolean
Public GetHashCodeResult As Long
Public ToStringResult As String


Private Function IObject_Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        Set Me.Equals = Value
    Else
        Me.Equals = Value
    End If
    
    IObject_Equals = Me.EqualsResult
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = Me.GetHashCodeResult
End Function

Private Function IObject_ToString() As String
    IObject_ToString = Me.ToStringResult
End Function
