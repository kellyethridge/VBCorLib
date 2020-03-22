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

Private mInited As Boolean
Private mValue As Long

Public Property Get Value() As Long
    Value = mValue
End Property

Public Sub Init(ByVal Value As Long)
    Debug.Assert Not mInited
    mValue = Value
    mInited = True
End Sub

Private Function IObject_Equals(Value As Variant) As Boolean
    On Error GoTo Catch
    Dim Other As Int32
    Set Other = Value
    IObject_Equals = Me.Value = Other.Value
    
Catch:
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = Me.Value
End Function

Private Function IObject_ToString() As String
    IObject_ToString = Value
End Function

Private Function IComparable_CompareTo(Value As Variant) As Long
    Dim OtherValue As Long
    If IsObject(Value) Then
        Dim OtherInt32 As Int32
        
        If Value Is Nothing Then
            IComparable_CompareTo = 1
            Exit Function
        End If
        
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

