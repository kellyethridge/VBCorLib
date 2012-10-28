VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStopWatch
'
Option Explicit

Public Sub TestNotRunning()
    Dim sw As New StopWatch
    Assert.IsFalse sw.IsRunning, "Should not be running."
End Sub

Public Sub TestStartRunning()
    Dim sw As New StopWatch
    sw.StartCount
    Assert.IsTrue sw.IsRunning, "Should be running."
    sw.StopCount
    Assert.IsFalse sw.IsRunning, "Should not be running."
End Sub

