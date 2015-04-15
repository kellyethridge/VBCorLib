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

Public Sub CopyTo_WithNullArray_ThrowsArgumentNullException()
    Dim SUT As New Queue
    SUT.Enqueue 1
    Dim NullArray() As Long
    
    On Error Resume Next
    SUT.CopyTo NullArray, 0
    
    AssertArgumentNullException Err, "DstArray"
End Sub

Public Sub CopyTo_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim SUT As New Queue
    SUT.Equals 1
    Dim DstArray(1 To 5) As Long
    
    On Error Resume Next
    SUT.CopyTo DstArray, 0
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub CopyTo_WithMultiDimDstArray_ThrowsArgumentException()
    Dim SUT As New Queue
    Dim DstArray(0 To 2, 0 To 2) As Long
    
    On Error Resume Next
    SUT.CopyTo DstArray, 0
    
    AssertArgumentException Err
End Sub

Public Sub CopyTo_WhenItemsWontFit_ThrowsArgumentException()
    Dim SUT As New Queue
    SUT.Enqueue 1
    SUT.Enqueue 2
    SUT.Enqueue 3
    SUT.Enqueue 4
    Dim DstArray(1 To 4) As Long
    
    On Error Resume Next
    SUT.CopyTo DstArray, 2
    
    AssertArgumentException Err
End Sub

Public Sub CopyTo_WhenCalled_CopiesItemsToArray()
    Dim SUT As New Queue
    SUT.Enqueue 1
    SUT.Enqueue 2
    SUT.Enqueue 3
    SUT.Enqueue 4
    Dim Actual(1 To 4) As Long
    
    SUT.CopyTo Actual, 1
    
    Assert.That Actual, Iz.EqualTo(NewLongs(1, 2, 3, 4)).AsCollection
End Sub

Public Sub CopyTo_WithCircleQueueWrapped_CopiesItemsToArray()
    ' this test requires knowledge of the implementation unfortunately.
    Dim i As Long
    Dim SUT As New Queue
    For i = 1 To 16
        SUT.Enqueue i
    Next
    For i = 1 To 8
        SUT.Dequeue
        SUT.Enqueue i
    Next
    Dim Actual(0 To 15) As Long
    
    SUT.CopyTo Actual, 0
    
    Dim Expected() As Long
    Expected = NewLongs(9, 10, 11, 12, 13, 14, 15, 16, 1, 2, 3, 4, 5, 6, 7, 8)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

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

Public Sub Peek_WithEmptyQueue_ThrowsInvalidOperationException()
    Dim SUT As New Queue
    
    On Error Resume Next
    SUT.Peek
    
    AssertInvalidOperationException Err
End Sub

Public Sub Peek_WhenCalled_ReturnsNextValueWithoutAlteringQueue()
    Dim SUT As New Queue
    SUT.Enqueue 1
    SUT.Enqueue 2
    
    Dim Actual As Variant
    Actual = SUT.Peek
    
    Assert.That Actual, Iz.EqualTo(1)
    Assert.That SUT.Count, Iz.EqualTo(2)
End Sub

Public Sub ToArray_WithEmptyQueue_ReturnsEmptyArray()
    Dim SUT As New Queue
    
    Dim Actual() As Variant
    Actual = SUT.ToArray
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub ToArray_WhenCalled_ReturnsArrayOfItemsInOrderOfEnqueue()
    Dim SUT As New Queue
    SUT.Enqueue 1
    SUT.Enqueue 3
    SUT.Enqueue 2
    
    Dim Actual() As Variant
    Actual = SUT.ToArray
    
    Assert.That Actual, Iz.EqualTo(Array(1, 3, 2))
End Sub

Public Sub QueueShouldBePersistable()
    Dim SUT As New Queue
    SUT.Enqueue 1
    SUT.Enqueue "hi"
    
    Dim Actual As Queue
    Set Actual = Persist(SUT)
    
    Assert.That Actual, Iz.EqualTo(Array(1, "hi"))
End Sub

Public Sub TestInitWithElements()
    Dim SUT As Queue
    Dim Expected As Variant
    Expected = Array(1, 2, 3, 4, 5)

    Set SUT = NewQueue(Expected)
    
    Assert.That SUT, Iz.EqualTo(Expected)
End Sub
