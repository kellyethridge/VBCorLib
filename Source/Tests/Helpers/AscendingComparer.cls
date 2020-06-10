VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AscendingComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CorLib.IComparer

Private Function IComparer_Compare(a As Variant, b As Variant) As Long
    If a < b Then
        IComparer_Compare = -1
    ElseIf a > b Then
        IComparer_Compare = 1
    End If
End Function
