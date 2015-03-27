VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopWatchTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' StopWatchTests
'
Option Explicit

Public Sub IsRunning_WhenStopped_ReturnsFalse()
    Dim sw As New StopWatch
    
    Dim Actual As Boolean
    Actual = sw.IsRunning
    
    Assert.That Actual, Iz.False
End Sub

Public Sub StartCount_WhenCalled_Starts()
    Dim sw As New StopWatch
    
    sw.StartCount
    
    Assert.That sw.IsRunning, Iz.True
End Sub

Public Sub StartNew_WhenCalled_ReturnsRunningStopWatch()
    Dim sw As StopWatch
    
    Set sw = StopWatch.StartNew
    
    Assert.That sw.IsRunning, Iz.True
End Sub

Public Sub Ellapsed_AfterRunning_ReturnsNonZero()
    Dim sw As New StopWatch
    sw.StartCount
    
    Dim Actual As TimeSpan
    Set Actual = sw.Elapsed
    
    Assert.That Actual.Ticks, Iz.GreaterThan(0)
End Sub
