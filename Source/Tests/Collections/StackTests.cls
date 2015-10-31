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
'
' StackTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

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
            
        Case "NewStack_WithElements_PushesElementsOnStack"
            Test.Use Array(1, 2, 3)
            Test.Use NewArrayList(Source:=Array(1, 2, 3))
            
            Dim Col As New Collection
            Col.Add 1
            Col.Add 2
            Col.Add 3
            Test.Use Col
            
    End Select
End Sub

Public Sub Ctor_WhenCalled_IsEmpty()
    Dim SUT As New Stack
    
    Assert.That SUT, Iz.Empty
End Sub

Public Sub Push_WhenCalled_AddsOneToStack()
    Dim SUT As New Stack
    
    SUT.Push 1
    
    Assert.That SUT, Iz.EqualTo(Array(1))
End Sub

Public Sub Pop_WhenCalled_ReturnsValuesInOppositeOrderOfPush()
    Dim SUT As New Stack
    
    SUT.Push 1
    SUT.Push 2
    
    Assert.That SUT.Count, Iz.EqualTo(2)
    Assert.That SUT.Pop, Iz.EqualTo(2)
    Assert.That SUT.Pop, Iz.EqualTo(1)
End Sub

Public Sub Pop_WhenStackIsEmpty_ThrowsInvalidOperationException()
    Dim SUT As New Stack
    
    On Error Resume Next
    SUT.Pop
    
    AssertInvalidOperationException Err
End Sub

Public Sub Contains_WithValue_ReturnsExpected(ByVal Value As Variant, ByVal Expected As Boolean)
    Dim SUT As New Stack
    SUT.Push 1
    SUT.Push 2
    SUT.Push "a"
    SUT.Push #1/1/2001#
    
    Dim Actual As Boolean
    Actual = SUT.Contains(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Peek_WhenCalled_ReturnsNextItemToPop()
    Dim SUT As New Stack
    SUT.Push 1
    SUT.Push 2
    
    Dim Actual As Variant
    Actual = SUT.Peek
    
    Assert.That Actual, Iz.EqualTo(2)
    Assert.That SUT.Count, Iz.EqualTo(2)
End Sub

Public Sub Peek_WithEmptyStack_ThrowsInvalidOperationException()
    Dim SUT As New Stack
    
    On Error Resume Next
    SUT.Peek
    
    AssertInvalidOperationException Err
End Sub

Public Sub ToArray_WhenCalled_ReturnsArrayOfElementsInOrderOfPop()
    Dim SUT As New Stack
    SUT.Push 1
    SUT.Push 2
    SUT.Push "a"
    
    Dim Actual() As Variant
    Actual = SUT.ToArray
    
    Assert.That Actual, Iz.EqualTo(Array("a", 2, 1))
End Sub

Public Sub ToArray_WithEmptyStack_ReturnsEmptyArray()
    Dim SUT As New Stack
    
    Dim Actual() As Variant
    Actual = SUT.ToArray
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub Clear_WhenCalled_EmptiesStack()
    Dim SUT As New Stack
    SUT.Push 1
    
    SUT.Clear
    
    Assert.That SUT, Iz.Empty
End Sub

Public Sub Enumeration_WhenCalled_ReturnsElementsInOrderOfPop()
    Dim SUT As New Stack
    SUT.Push 1
    SUT.Push 2
    SUT.Push 3
    
    Dim Actual(0 To 2) As Variant
    Dim Index As Long
    Dim Item As Variant
    For Each Item In SUT
        Actual(Index) = Item
        Index = Index + 1
    Next
    
    Assert.That Actual, Iz.EqualTo(Array(3, 2, 1))
End Sub

Public Sub CopyTo_WhenCalled_CopiesItemsToArrayInOrderOfPop()
    Dim SUT As New Stack
    SUT.Push 1
    SUT.Push 2
    
    Dim Actual(0 To 1) As Variant
    SUT.CopyTo Actual, 0
    
    Assert.That Actual, Iz.EqualTo(Array(2, 1))
End Sub

Public Sub CopyTo_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim SUT As New Stack
    SUT.Push 1
    
    On Error Resume Next
    Dim Target(0 To 1) As Variant
    SUT.CopyTo Target, -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub CopyTo_WithMultiDimensionalArray_ThrowsArgumentException()
    Dim SUT As New Stack
    SUT.Push 1
    
    On Error Resume Next
    Dim Target(0 To 1, 0 To 1) As Variant
    SUT.CopyTo Target, 0
    
    AssertArgumentException Err, "Arr"
End Sub

Public Sub CopyTo_WithNotEnoughRoomToCopyStack_ThrowsArgumentException()
    Dim SUT As New Stack
    SUT.Push 1
    SUT.Push 2
    SUT.Push 3
    
    On Error Resume Next
    Dim Target(0 To 2)
    SUT.CopyTo Target, 1
    
    AssertArgumentException Err
End Sub

Public Sub Serialization_WhenCalled_SerializesInstance()
    Dim SUT As New Stack
    SUT.Push 1
    
    Dim Actual As Stack
    Set Actual = Persist(SUT)
    
    Assert.That Actual.Pop, Iz.EqualTo(1)
End Sub

Public Sub Push_WithLargeAmountOfData_StoresOnStack()
    Dim SUT As New Stack
    Dim i As Long
    
    For i = 1 To 100
        SUT.Push i
    Next
    
    For i = 100 To 1 Step -1
        Assert.That SUT.Pop, Iz.EqualTo(i)
    Next
End Sub

Public Sub NewStack_WithElements_PushesElementsOnStack(ByVal Source As Variant)
    Dim SUT As Stack
    
    Set SUT = NewStack(Source)
    
    Assert.That SUT.ToArray, Iz.EqualTo(Array(3, 2, 1))
End Sub

Public Sub NewStack_WithAnArray_DoesNotAlterArray()
    Dim SUT As Stack
    Dim Source As Variant
    Source = Array(1, 2, 3)
    
    NewStack Source
    
    Assert.That Source, Iz.EqualTo(Array(1, 2, 3))
End Sub

Public Sub Clone_WhenCalled_ReturnsCopyOfStack()
    Dim SUT As New Stack
    SUT.Push 1
    
    Dim Actual As Stack
    Set Actual = SUT.Clone
    
    Assert.That Actual.Pop, Iz.EqualTo(1)
    Assert.That SUT.Pop, Iz.EqualTo(1)
End Sub
