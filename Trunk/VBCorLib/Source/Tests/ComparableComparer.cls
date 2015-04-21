VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComparableComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements SimplyVBUnitType.IComparer

Private Function IComparer_Compare(x As Variant, y As Variant) As Long
    Dim XComparable As IComparable
    Set XComparable = x
    
    IComparer_Compare = XComparable.CompareTo(y)
End Function
