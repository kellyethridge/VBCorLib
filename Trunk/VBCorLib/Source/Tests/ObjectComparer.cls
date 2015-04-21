VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements SimplyVBUnitType.IEqualityComparer

Private Function IEqualityComparer_Equals(X As Variant, Y As Variant) As Boolean
    Dim Left    As IObject
    Dim Right   As IObject
    
    Set Left = X
    Set Right = Y
    
    IEqualityComparer_Equals = Left.Equals(Right)
End Function
