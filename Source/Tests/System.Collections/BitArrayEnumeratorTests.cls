VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitArrayEnumeratorTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BitArrayEnumeratorTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Public Sub Current_WhenNotMovedNext_ThrowsInvalidOperationException()
    Dim En As IEnumerator
    Set En = MakeEnumerator
    
    On Error Resume Next
    Dim Value As Variant
    Value = En.Current
    
    AssertInvalidOperationException Err
End Sub

Public Sub Current_WhenMoveNextIsCalled_ReturnsItem()
    Dim En As IEnumerator
    Set En = MakeEnumerator
    
    Assert.That En.MoveNext, Iz.True
    Dim Actual As Variant
    Actual = En.Current
    
    Assert.That Actual, Iz.EqualTo(True)
End Sub

Public Sub MoveNext_WhenMovingPassedEnd_ReturnsFalse()
    Dim En As IEnumerator
    Set En = MakeEnumerator
    
    En.MoveNext
    Dim Actual As Boolean
    Actual = En.MoveNext
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Current_WhenMoveNextReturnsFalse_ThrowsInvalidOperationException()
    Dim En As IEnumerator
    Set En = MakeEnumerator
    En.MoveNext
    En.MoveNext
    
    On Error Resume Next
    Dim Actual As Variant
    Actual = En.Current
    
    AssertInvalidOperationException Err
End Sub

Public Sub MoveNext_AfterHavingPassedEnd_ThrowsInvalidOperationException()
    Dim En As IEnumerator
    Set En = MakeEnumerator
    En.MoveNext
    En.MoveNext
    
    On Error Resume Next
    En.MoveNext
    
    AssertInvalidOperationException Err
End Sub

Public Sub Current_WhenListHasChanged_ThrowsInvalidOperationException()
    Dim List As New BitArray
    List(1) = True
    Dim En As IEnumerator
    Set En = List.GetEnumerator
    En.MoveNext
    List(2) = True
    
    On Error Resume Next
    Dim Actual As Variant
    Actual = En.Current
    
    AssertInvalidOperationException Err
End Sub

Public Sub Reset_WhenCalled_RestartsEnumeration()
    Dim En As IEnumerator
    Set En = MakeEnumerator
    En.MoveNext
    En.MoveNext
    En.Reset
    En.MoveNext
    
    Dim Actual As Variant
    Actual = En.Current
    
    Assert.That Actual, Iz.EqualTo(True)
End Sub

Private Function MakeEnumerator() As IEnumerator
    Dim List As BitArray
    Set List = Cor.NewBitArray(1, True)
    
    Set MakeEnumerator = List.GetEnumerator
End Function

