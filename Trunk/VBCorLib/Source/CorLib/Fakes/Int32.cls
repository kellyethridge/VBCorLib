VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Int32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IComparable
Implements IObject

Public Value As Long

Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = object.Equals(Me, Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = ObjPtr(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = Value
End Function

Private Function IComparable_CompareTo(Value As Variant) As Long
    Dim OtherValue As Long
    If IsObject(Value) Then
        Dim OtherInt32 As Int32
        Set OtherInt32 = Value
        OtherValue = OtherInt32.Value
    Else
        OtherValue = Value
    End If
    
    If Me.Value < OtherValue Then
        IComparable_CompareTo = -1
    ElseIf Me.Value > OtherValue Then
        IComparable_CompareTo = 1
    End If
End Function

