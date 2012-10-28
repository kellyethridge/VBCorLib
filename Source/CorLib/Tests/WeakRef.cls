VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeakRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' WeakRef
'
Option Explicit

Private mStrong As StrongRef


Public Property Get Reference() As StrongRef
    Set Reference = mStrong
End Property
Public Property Set Reference(ByVal RHS As StrongRef)
    Set mStrong = RHS
End Property
