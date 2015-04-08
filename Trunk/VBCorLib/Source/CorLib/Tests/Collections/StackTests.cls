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
            
        Case "NewStack_WithElements_PushesElementsOnStack"
            Test.Use Array(1, 2, 3)
            Test.Use NewArrayList(c:=Array(1, 2, 3))
            
            Dim Col As New Collection
            Col.Add 1
            Col.Add 2
            Col.Add 3
            Test.Use Col
            
    End Select
End Sub

Public Sub Ctor_WhenCalled_IsEmpty()
    Dim Sut As New Stack
    
    Assert.That Sut, Iz.Empty
End Sub

Public Sub Push_WhenCalled_AddsOneToStack()
    Dim Sut As New Stack
    
    Sut.Push 1
    
    Assert.That Sut, Iz.EqualTo(Array(1))
End Sub

Public Sub Pop_WhenCalled_ReturnsValuesInOppositeOrderOfPush()
    Dim Sut As New Stack
    
    Sut.Push 1
    Sut.Push 2
    
    Assert.That Sut.Count, Iz.EqualTo(2)
    Assert.That Sut.Pop, Iz.EqualTo(2)
    Assert.That Sut.Pop, Iz.EqualTo(1)
End Sub

Public Sub Pop_WhenStackIsEmpty_ThrowsInvalidOperationException()
    Dim Sut As New Stack
    
    On Error GoTo errTrap
    Sut.Pop
    
errTrap:
    AssertInvalidOperationException Err
End Sub

Public Sub Contains_WithValue_ReturnsExpected(ByVal Value As Variant, ByVal Expected As Boolean)
    Dim Sut As New Stack
    Sut.Push 1
    Sut.Push 2
    Sut.Push "a"
    Sut.Push #1/1/2001#
    
    Dim Actual As Boolean
    Actual = Sut.Contains(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Peek_WhenCalled_ReturnsNextItemToPop()
    Dim Sut As New Stack
    Sut.Push 1
    Sut.Push 2
    
    Dim Actual As Variant
    Actual = Sut.Peek
    
    Assert.That Actual, Iz.EqualTo(2)
    Assert.That Sut.Count, Iz.EqualTo(2)
End Sub

Public Sub Peek_WithEmptyStack_ThrowsInvalidOperationException()
    Dim Sut As New Stack
    
    On Error GoTo errTrap
    Sut.Peek
    
errTrap:
    AssertInvalidOperationException Err
End Sub

Public Sub ToArray_WhenCalled_ReturnsArrayOfElementsInOrderOfPop()
    Dim Sut As New Stack
    Sut.Push 1
    Sut.Push 2
    Sut.Push "a"
    
    Dim Actual() As Variant
    Actual = Sut.ToArray
    
    Assert.That Actual, Iz.EqualTo(Array("a", 2, 1))
End Sub

Public Sub ToArray_WithEmptyStack_ReturnsEmptyArray()
    Dim Sut As New Stack
    
    Dim Actual() As Variant
    Actual = Sut.ToArray
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub Clear_WhenCalled_EmptiesStack()
    Dim Sut As New Stack
    Sut.Push 1
    
    Sut.Clear
    
    Assert.That Sut, Iz.Empty
End Sub

Public Sub Enumeration_WhenCalled_ReturnsElementsInOrderOfPop()
    Dim Sut As New Stack
    Sut.Push 1
    Sut.Push 2
    Sut.Push 3
    
    Dim Actual(0 To 2) As Variant
    Dim Index As Long
    Dim Item As Variant
    For Each Item In Sut
        Actual(Index) = Item
        Index = Index + 1
    Next
    
    Assert.That Actual, Iz.EqualTo(Array(3, 2, 1))
End Sub

Public Sub CopyTo_WhenCalled_CopiesItemsToArrayInOrderOfPop()
    Dim Sut As New Stack
    Sut.Push 1
    Sut.Push 2
    
    Dim Actual(0 To 1) As Variant
    Sut.CopyTo Actual, 0
    
    Assert.That Actual, Iz.EqualTo(Array(2, 1))
End Sub

Public Sub CopyTo_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Sut As New Stack
    Sut.Push 1
    
    On Error Resume Next
    Dim Target(0 To 1) As Variant
    Sut.CopyTo Target, -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub CopyTo_WithMultiDimensionalArray_ThrowsArgumentException()
    Dim Sut As New Stack
    Sut.Push 1
    
    On Error Resume Next
    Dim Target(0 To 1, 0 To 1) As Variant
    Sut.CopyTo Target, 0
    
    AssertArgumentException Err, "DstArray"
End Sub

Public Sub CopyTo_WithNotEnoughRoomToCopyStack_ThrowsArgumentException()
    Dim Sut As New Stack
    Sut.Push 1
    Sut.Push 2
    Sut.Push 3
    
    On Error Resume Next
    Dim Target(0 To 2)
    Sut.CopyTo Target, 1
    
    AssertArgumentException Err
End Sub

Public Sub TestPersitence()
    Dim Sut As New Stack
    Sut.Push 1
    
    Dim Bag As New PropertyBag
    Bag.WriteProperty "Stack", Sut
    
    Dim Actual As Stack
    Set Actual = Bag.ReadProperty("Stack")
    Assert.That Actual.Pop, Iz.EqualTo(1)
End Sub

Public Sub Push_WithLargeAmountOfData_StoresOnStack()
    Dim Sut As New Stack
    Dim i As Long
    
    For i = 1 To 100
        Sut.Push i
    Next
    
    For i = 100 To 1 Step -1
        Assert.That Sut.Pop, Iz.EqualTo(i)
    Next
End Sub

Public Sub NewStack_WithElements_PushesElementsOnStack(ByVal Source As Variant)
    Dim Sut As Stack
    
    Set Sut = NewStack(Source)
    
    Assert.That Sut.ToArray, Iz.EqualTo(Array(3, 2, 1))
End Sub

Public Sub NewStack_WithAnArray_DoesNotAlterArray()
    Dim Sut As Stack
    Dim Source As Variant
    Source = Array(1, 2, 3)
    
    NewStack Source
    
    Assert.That Source, Iz.EqualTo(Array(1, 2, 3))
End Sub
