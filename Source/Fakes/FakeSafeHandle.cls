VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeSafeHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements SafeHandle

Private WithEvents mBase As SafeHandle
Attribute mBase.VB_VarHelpID = -1

Public ReleaseHandleCallCount As New Count


Public Sub Init(ByVal Handle As Long, ByVal OwnsHandle As Boolean)
    Set mBase = Cor.NewSafeHandle(Handle, OwnsHandle)
End Sub

Public Sub CloseHandle()
    mBase.CloseHandle
End Sub

Public Function DangerousGetHandle() As Long
    DangerousGetHandle = mBase.DangerousGetHandle
End Function

Public Property Get IsClosed() As Boolean
    IsClosed = mBase.IsClosed
End Property

Public Property Get IsInvalid() As Boolean
    IsInvalid = mBase.IsInvalid
End Property

Private Sub Class_Terminate()
    CloseHandle
End Sub

Private Sub mBase_ReleaseHandle()
    ReleaseHandleCallCount.Increment
End Sub

Private Sub SafeHandle_CloseHandle()
    CloseHandle
End Sub

Private Function SafeHandle_DangerousGetHandle() As Long
    SafeHandle_DangerousGetHandle = DangerousGetHandle
End Function

Private Property Get SafeHandle_IsClosed() As Boolean
    SafeHandle_IsClosed = IsClosed
End Property

Private Property Get SafeHandle_IsInvalid() As Boolean
    SafeHandle_IsInvalid = IsInvalid
End Property
