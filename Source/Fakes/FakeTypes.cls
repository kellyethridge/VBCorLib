VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeTypes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type LargeUdt
    Value As Long
    Large1 As Variant
    Large2 As Variant
End Type


Public Function NewLargeUdt(ByVal Value As Long, Optional ByRef Large1 As Variant, Optional ByRef Large2 As Variant) As LargeUdt
    NewLargeUdt.Value = Value
    NewLargeUdt.Large1 = Large1
    NewLargeUdt.Large2 = Large2
End Function
