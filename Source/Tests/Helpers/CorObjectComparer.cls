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

Private mComparer As CorLib.IComparer


Private Sub Class_Initialize()
    Set mComparer = Comparer.Default
End Sub

Private Function IComparer_Compare(x As Variant, y As Variant) As Long
    IComparer_Compare = mComparer.Compare(x, y)
End Function

Private Function IEqualityComparer_Equals(x As Variant, y As Variant) As Boolean
    IEqualityComparer_Equals = Object.Equals(x, y)
End Function

Friend Sub Init(ByVal Comparer As CorLib.IComparer)
    Set mComparer = Comparer
End Sub
