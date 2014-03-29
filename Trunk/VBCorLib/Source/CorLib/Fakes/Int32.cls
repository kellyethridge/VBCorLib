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
'
'   Int32
'
'   This is a simple class to represent an object version of a 32bit number.
'
Option Explicit
Implements IComparable
Implements IObject

Public mValue As Long


Private Sub Class_Terminate()
    'Debug.Print "Terminated"
End Sub

Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = object.Equals(Me, Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = ObjPtr(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = mValue
End Function

Private Function IComparable_CompareTo(Value As Variant) As Long
    Dim l As Long
    If IsObject(Value) Then
        Dim o As Int32
        Set o = Value
        l = o.mValue
    Else
        l = Value
    End If
    
    If mValue < l Then IComparable_CompareTo = -1: Exit Function
    If mValue > l Then IComparable_CompareTo = 1
End Function

