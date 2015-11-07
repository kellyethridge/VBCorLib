VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashtableValueCollectionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' HashtableValueCollectionTests
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
    
    Assert.That Actual, Iz.EqualTo(1)
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

Public Sub Current_WhenTableHasChanged_ThrowsInvalidOperationException()
    Dim Table As New Hashtable
    Table.Add "a", 1
    Dim En As IEnumerator
    Set En = Table.GetEnumerator
    En.MoveNext
    Table.Add "b", 2
    
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
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub CopyTo_WhenCalled_CopiesEntriesToArray()
    Dim Table As New Hashtable
    Table.Add "a", 1
    Table.Add "b", 2
    Dim Target(0 To 1) As Long
    
    Table.Values.CopyTo Target, 0
    
    Assert.That Target, Iz.EquivalentTo(Array(1, 2))
End Sub

Public Sub CopyTo_WithIndex_CopiesEntriesToArrayOffset()
    Dim Table As New Hashtable
    Table.Add "a", 1
    Table.Add "b", 2
    Dim Target(0 To 2) As Variant
    
    Table.Values.CopyTo Target, 1
    
    Assert.That Target(0), Iz.vbEmpty
    Assert.That Target, Iz.EquivalentTo(Array(Empty, 1, 2))
End Sub

Public Sub ForEach_WhenCalled_EnumeratesKeys()
    Dim Table As New Hashtable
    Table.Add "a", 1
    Table.Add "b", 2
    
    Dim Values As New ArrayList
    Dim Value As Variant
    For Each Value In Table.Values
        Values.Add Value
    Next
    
    Assert.That Values, Iz.EquivalentTo(Array(1, 2))
End Sub


Private Function MakeEnumerator() As IEnumerator
    Dim Table As New Hashtable
    Table.Add "a", 1
    
    Set MakeEnumerator = Table.Values.GetEnumerator
End Function



