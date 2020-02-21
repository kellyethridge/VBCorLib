VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorObjectComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements SimplyVBUnitType.IEqualityComparer
Implements SimplyVBUnitType.IComparer

Private Function IComparer_Compare(X As Variant, Y As Variant) As Long
    IComparer_Compare = Comparer.Default.Compare(X, Y)
End Function

Private Function IEqualityComparer_Equals(X As Variant, Y As Variant) As Boolean
    IEqualityComparer_Equals = Object.Equals(X, Y)
End Function
