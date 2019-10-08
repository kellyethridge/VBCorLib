VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LargeUdtComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements SimplyVBUnitType.IComparer

Private Function IComparer_Compare(a As Variant, b As Variant) As Long
    Dim x As LargeUdt
    Dim y As LargeUdt
    
    x = a
    y = b
    
    IComparer_Compare = Sgn(x.Value - y.Value)
End Function
