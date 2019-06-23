VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StrongRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' StrongRef
'
Option Explicit

Private mWeak As WeakReference


Public Property Get Reference() As WeakRef
    Set Reference = mWeak.Target
End Property

Public Property Set Reference(ByVal RHS As WeakRef)
    Set mWeak.Target = RHS
End Property

Public Sub AssertDead()
    Assert.IsFalse mWeak.IsAlive, "WeakReference should be dead."
End Sub

Public Sub AssertAlive()
    Assert.IsTrue mWeak.IsAlive, "WeakReference should be alive."
End Sub

Private Sub Class_Initialize()
    Set mWeak = New WeakReference
End Sub

Private Sub Class_Terminate()
    Debug.Print "StrongRef: Terminate"
End Sub
