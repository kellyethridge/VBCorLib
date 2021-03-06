VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashtableEnumeratorTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' HashtableEnumeratorTests
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
    Dim Value As DictionaryEntry
    Set Value = En.Current
    
    AssertInvalidOperationException Err
End Sub

Public Sub Current_WhenMoveNextIsCalled_ReturnsItem()
    Dim En As IEnumerator
    Set En = MakeEnumerator
    
    Assert.That En.MoveNext, Iz.True
    Dim Actual As DictionaryEntry
    Set Actual = En.Current
    
    Assert.That Actual.Key, Iz.EqualTo("a")
    Assert.That Actual.Value, Iz.EqualTo(1)
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
    Dim Actual As DictionaryEntry
    Set Actual = En.Current
    
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
    Dim Actual As DictionaryEntry
    Set Actual = En.Current
    
    AssertInvalidOperationException Err
End Sub

Public Sub Reset_WhenCalled_RestartsEnumeration()
    Dim En As IEnumerator
    Set En = MakeEnumerator
    En.MoveNext
    En.MoveNext
    En.Reset
    En.MoveNext
    
    Dim Actual As DictionaryEntry
    Set Actual = En.Current
    
    Assert.That Actual.Key, Iz.EqualTo("a")
    Assert.That Actual.Value, Iz.EqualTo(1)
End Sub

Public Sub ForEach_WhenCalled_EnumeratesHashtable()
    Dim Table As New Hashtable
    Dim Keys As New Collection
    Dim Values As New Collection
    Dim Entry As DictionaryEntry
    
    Table.Add "a", 1
    
    For Each Entry In Table
        Keys.Add Entry.Key
        Values.Add Entry.Value
    Next
    
    Assert.That Keys, Iz.EquivalentTo(Array("a"))
    Assert.That Values, Iz.EquivalentTo(Array(1))
End Sub

Public Sub ForEach_WhenCalledAsDictionary_EnumeratesDictionary()
    Dim Table As New Hashtable
    Dim Keys As New Collection
    Dim Values As New Collection
    Dim Entry As DictionaryEntry
    
    Table.Add "a", 1
    
    For Each Entry In AsDictionary(Table)
        Keys.Add Entry.Key
        Values.Add Entry.Value
    Next
    
    Assert.That Keys, Iz.EquivalentTo(Array("a"))
    Assert.That Values, Iz.EquivalentTo(Array(1))
End Sub

Private Function AsDictionary(ByVal Table As Hashtable) As IDictionary
    Set AsDictionary = Table
End Function

Private Function MakeEnumerator() As IEnumerator
    Dim Table As New Hashtable
    Table.Add "a", 1
    
    Set MakeEnumerator = Table.GetEnumerator
End Function

