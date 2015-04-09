VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueueTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Const DEF_CAPACITY As Long = 16

Private Q As Queue

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Contains_WithValues_ReturnsExpected"
            Test.Use 1, True
            Test.Use "1", False
            Test.Use "a", True
            Test.Use CObj(Me), True
            Test.Use 2, False
        
    End Select
End Sub


Public Sub Setup()
    Set Q = New Queue
End Sub

Public Sub Ctor_WhenCalled_CreatesEmptyQueue()
    Dim SUT As Queue
    
    Set SUT = New Queue
    
    Assert.That SUT, Iz.Empty
    Assert.That SUT.Count, Iz.EqualTo(0)
End Sub

Public Sub Enqueue_WhenCalled_UpdatesCount()
    Dim SUT As New Queue
    
    SUT.Enqueue 1
    SUT.Enqueue 2
    
    Assert.That SUT.Count, Iz.EqualTo(2)
End Sub

Public Sub Enqueue_WhenCalled_AddsValueFifoOrder()
    Dim SUT As New Queue
    
    SUT.Enqueue 1
    SUT.Enqueue 2
    
    Assert.That SUT.Dequeue, Iz.EqualTo(1)
    Assert.That SUT.Dequeue, Iz.EqualTo(2)
    Assert.That SUT, Iz.Empty
End Sub

Public Sub Enqueue_WithMoreThanDefaultCapacity_ExtendsCapacity()
    Const NumberAdded   As Long = DEF_CAPACITY * 2
    Dim SUT             As New Queue
    Dim i               As Long
        
    For i = 1 To NumberAdded
        SUT.Enqueue i
    Next
    
    Assert.That SUT.Count, Iz.EqualTo(NumberAdded)
    For i = 1 To NumberAdded
        Assert.That SUT.Dequeue, Iz.EqualTo(i)
    Next
    Assert.That SUT, Iz.Empty
End Sub

Public Sub Clear_WhenCalled_EmptiesQueue()
    Dim SUT As New Queue
    SUT.Enqueue 1
    SUT.Enqueue 2
    
    SUT.Clear
    
    Assert.That SUT, Iz.Empty
End Sub

Public Sub Clone_WhenCalled_ReturnsCopyOfQueue()
    Dim SUT As New Queue
    SUT.Enqueue 1
    SUT.Enqueue "b"
    
    Dim SUTCopy As Queue
    Set SUTCopy = SUT.Clone
    
    Assert.That SUTCopy, Iz.Not.SameAs(SUT)
    Assert.That SUTCopy.Dequeue, Iz.EqualTo(1)
    Assert.That SUTCopy.Dequeue, Iz.EqualTo("b")
    Assert.That SUT.Dequeue, Iz.EqualTo(1)
    Assert.That SUT.Dequeue, Iz.EqualTo("b")
End Sub

Public Sub Contains_WithValues_ReturnsExpected(ByVal Value As Variant, ByVal Expected As Boolean)
    Dim SUT As New Queue
    SUT.Enqueue 1
    SUT.Enqueue "a"
    SUT.Enqueue CObj(Me)
    
    Dim Actual As Boolean
    Actual = SUT.Contains(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'Public Sub TestContainsWithComparer()
'    Dim i As Long
'    For i = 1 To 100
'        Q.Enqueue i
'    Next i
'
'    Assert.IsTrue Q.Contains(99&, Comparer.Default), "q should contain 99."
'End Sub
'
'Public Sub TestCopyTo()
'    Dim i As Long
'    For i = 1 To 100
'        Q.Enqueue i
'    Next i
'
'    Dim a(99) As Long
'    Q.CopyTo a, 0
'    For i = 1 To 100
'        Assert.AreEqual i, a(i - 1)
'    Next i
'End Sub
'
'Public Sub TestCopyToWrapping()
'    ' required INDEBUG=1 VBCorLib compiler constant.
'    Dim i As Long
'    For i = 1 To 8
'        Q.Enqueue i
'        Q.Dequeue
'    Next i
'    For i = 1 To 14
'        Q.Enqueue i
'    Next i
'
'    Dim a(-5 To 8) As Long
'    Q.CopyTo a, -5
'    For i = -5 To 8
'        Assert.AreEqual i + 6, a(i)
'    Next i
'End Sub
'
'Public Sub TestEnumerator()
'    Dim i As Long
'    For i = 1 To 100
'        Q.Enqueue i
'    Next i
'    Dim v As Variant
'    i = 1
'    For Each v In Q
'        Assert.AreEqual i, v
'        i = i + 1
'    Next v
'End Sub
'
'Public Sub TestPeek()
'    Q.Enqueue 1
'    Assert.AreEqual 1, Q.Peek, "Peek should return 1."
'    Assert.AreEqual 1, Q.Count, "Count should not change."
'End Sub
'
'Public Sub TestToArray()
'    Dim i As Long
'    For i = 1 To 10
'        Q.Enqueue 1
'        Q.Dequeue
'    Next i
'    For i = 11 To 20
'        Q.Enqueue i
'    Next i
'    Dim a() As Variant
'    a = Q.ToArray
'    Assert.AreEqual Q.Count, cArray.GetLength(a), "Wrong array length."
'
'    For i = 11 To 20
'        Assert.AreEqual i, a(i - 11)
'    Next i
'End Sub
'
'Public Sub TestEmptyToArray()
'    Dim a() As Variant
'    a = Q.ToArray
'    Assert.IsFalse cArray.IsNull(a), "Array should not be null."
'    Assert.AreEqual 0, cArray.GetLength(a), "Array should be zero length."
'End Sub
'
'Public Sub TestTrimToSize()
'    ' requires INDEBUG=1 VBCorLib compiler constant.
'    Dim i As Long
'    For i = 1 To 8
'        Q.Enqueue i
'    Next i
'    Q.Dequeue
'    Q.TrimToSize
'
'    For i = 2 To 8
'        Assert.AreEqual i, Q.Dequeue, "Queue contents should not change."
'    Next i
'End Sub
'
'Public Sub TestTrimToSize2()
'    Dim i As Long
'    For i = 1 To 16
'        Q.Enqueue i
'    Next i
'    For i = 1 To 5
'        Q.Dequeue
'    Next i
'    For i = 17 To 21
'        Q.Enqueue i
'    Next i
'
'    Q.TrimToSize
'
'    For i = 6 To 21
'        Assert.AreEqual i, Q.Dequeue
'    Next i
'End Sub
'
'Public Sub TestPersistence()
'    Dim i As Long
'    For i = 1 To 100
'        Q.Enqueue i
'    Next i
'    Dim pp1 As New PropertyBag
'    Dim pp2 As New PropertyBag
'    Dim q2 As Queue
'
'    pp1.WriteProperty "q", Q
'    pp2.Contents = pp1.Contents
'    Set q2 = pp2.ReadProperty("q")
'
'    For i = 1 To 100
'        Assert.AreEqual i, q2.Dequeue
'    Next i
'End Sub
'
'Public Sub TestInitWithElements()
'    Dim a() As Long
'    a = NewArray(vbLong, 1, 2, 3, 4, 5)
'
'    Set Q = NewQueue(a)
'    Assert.AreEqual 5, Q.Count
'
'    Dim i As Long
'    For i = 1 To 5
'        Assert.AreEqual i, Q.Dequeue
'    Next i
'End Sub
