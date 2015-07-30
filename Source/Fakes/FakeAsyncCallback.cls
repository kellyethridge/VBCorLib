VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeAsyncCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements AsyncCallback

Private mResult As IAsyncResult
Private mCalled As Boolean

Public Property Get Result() As IAsyncResult
    Set Result = mResult
End Property

Public Property Get Called() As Boolean
    Called = mCalled
End Property

Private Sub AsyncCallback_Execute(ByVal ar As IAsyncResult)
    mCalled = True
End Sub
