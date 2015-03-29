VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StackTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Contains_WithValue_ReturnsExpected"
            Test.Use 1, True
            Test.Use 2, True
            Test.Use 3, False
            Test.Use "a", True
            Test.Use "b", False
            Test.Use #1/1/2002#, False
            Test.Use #1/1/2001#, True
            
    End Select
End Sub

Public Sub Ctor_WhenCalled_IsEmpty()
    Dim sut As New Stack
    
    Assert.That sut, Iz.Empty
End Sub

Public Sub Push_WhenCalled_AddsOneToStack()
    Dim sut As New Stack
    
    sut.Push 1
    
    Assert.That sut, Iz.EqualTo(Array(1))
End Sub

Public Sub Pop_WhenCalled_ReturnsValuesInOppositeOrderOfPush()
    Dim sut As New Stack
    
    sut.Push 1
    sut.Push 2
    
    Assert.That sut.Count, Iz.EqualTo(2)
    Assert.That sut.Pop, Iz.EqualTo(2)
    Assert.That sut.Pop, Iz.EqualTo(1)
End Sub

Public Sub Pop_WhenStackIsEmpty_ThrowsInvalidOperationException()
    Dim sut As New Stack
    
    On Error GoTo errTrap
    sut.Pop
    
errTrap:
    AssertInvalidOperationException Err
End Sub

Public Sub Contains_WithValue_ReturnsExpected(ByVal Value As Variant, ByVal Expected As Boolean)
    Dim sut As New Stack
    sut.Push 1
    sut.Push 2
    sut.Push "a"
    sut.Push #1/1/2001#
    
    Dim Actual As Boolean
    Actual = sut.Contains(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Peek_WhenCalled_ReturnsNextItemToPop()
    Dim sut As New Stack
    sut.Push 1
    sut.Push 2
    
    Dim Actual As Variant
    Actual = sut.Peek
    
    Assert.That Actual, Iz.EqualTo(2)
    Assert.That sut.Count, Iz.EqualTo(2)
End Sub

Public Sub ToArray_WhenCalled_ReturnsArrayOfElementsInOrderOfPop()
    Dim sut As New Stack
    sut.Push 1
    sut.Push 2
    sut.Push "a"
    
    Dim Actual() As Variant
    Actual = sut.ToArray
    
    Assert.That Actual, Iz.EqualTo(Array("a", 2, 1))
End Sub

Public Sub ToArray_WithEmptyStack_ReturnsEmptyArray()
    Dim sut As New Stack
    
    Dim Actual() As Variant
    Actual = sut.ToArray
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub Clear_WhenCalled_EmptiesStack()
    Dim sut As New Stack
    sut.Push 1
    
    sut.Clear
    
    Assert.That sut, Iz.Empty
End Sub

Public Sub Enumeration_WhenCalled_ReturnsElementsInOrderOfPop()
    Dim sut As New Stack
    sut.Push 1
    sut.Push 2
    sut.Push 3
    
    Dim Actual(0 To 2) As Variant
    Dim Index As Long
    Dim Item As Variant
    For Each Item In sut
        Actual(Index) = Item
        Index = Index + 1
    Next
    
    Assert.That Actual, Iz.EqualTo(Array(3, 2, 1))
End Sub

'Public Sub TestCopyTo()
'    Dim i As Long
'    For i = 1 To 10
'        st.Push i
'    Next i
'
'    Dim a(0 To 9) As Long
'    st.CopyTo a, 0
'    For i = 0 To 9
'        Assert.AreEqual 10 - i, a(i)
'    Next i
'End Sub

'Public Sub TestPersitence()
'    Dim i As Long
'    For i = 1 To 10
'        st.Push i
'    Next i
'
'    Dim Copy As Stack
'    Dim pp1 As New PropertyBag
'    Dim pp2 As New PropertyBag
'    pp1.WriteProperty "stack", st
'    pp2.Contents = pp1.Contents
'
'    Set Copy = pp2.ReadProperty("stack", Nothing)
'    For i = 1 To 10
'        Assert.AreEqual st.Pop, Copy.Pop
'    Next i
'End Sub
'
'Public Sub TestBigStack()
'    Dim i As Long
'    For i = 1 To 100
'        st.Push i
'    Next i
'
'    For i = 100 To 1 Step -1
'        Assert.AreEqual i, st.Pop
'    Next i
'End Sub
'
'Public Sub TestInitWithElements()
'    Dim a() As Long
'    a = NewArray(vbLong, 1, 2, 3, 4, 5)
'
'    Set st = NewStack(a)
'    Assert.AreEqual 5, st.Count
'
'    Dim i As Long
'    For i = 1 To 5
'        Assert.AreEqual 6 - i, st.Pop
'    Next i
'
'End Sub
