VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayListTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "LastIndexOf_WithValue_ReturnsExpected"
            Test.Use 1, 5
            Test.Use 5, 9
            
        Case "RemoveAt_WithValues_RemovesExpectedItem"
            Test.Use 0, NewLongs(2, 3, 4, 5, 1, 2, 3, 4, 5)
            Test.Use 9, NewLongs(1, 2, 3, 4, 5, 1, 2, 3, 4)
            Test.Use 5, NewLongs(1, 2, 3, 4, 5, 2, 3, 4, 5)
    End Select
End Sub

Public Sub Constructor_WhenCalled_CreatesEmptyList()
    Dim List As ArrayList
    
    Set List = New ArrayList
    
    Assert.That List.Count, Iz.EqualTo(0)
End Sub

Public Sub Add_WithValue_IncreasesCount()
    Dim List As New ArrayList
    
    List.Add 1
    
    Assert.That List.Count, Iz.EqualTo(1)
End Sub

Public Sub Add_WithValue_AppendsToList()
    Dim List As New ArrayList
    
    List.Add 1
    
    Assert.That List.Item(0), Iz.EqualTo(1)
End Sub

Public Sub Add_WithMultipleValues_AppendsToList()
    Dim List As New ArrayList
    
    List.Add 1
    List.Add 2
    
    Assert.That List(0), Iz.EqualTo(1)
    Assert.That List(1), Iz.EqualTo(2)
End Sub

Public Sub AddRange_WithCollection_AppendsToList()
    Dim Source As New Collection
    Source.Add 1
    Source.Add 2
    Source.Add 3
    Dim List As New ArrayList
    
    List.AddRange Source
    
    Assert.That List.Count, Iz.EqualTo(3)
    Assert.That List(0), Iz.EqualTo(1)
    Assert.That List(1), Iz.EqualTo(2)
    Assert.That List(2), Iz.EqualTo(3)
End Sub

Public Sub AddRange_WithArray_AppendsToList()
    Dim Source() As Long
    Source = NewLongs(1, 2, 3)
    Dim List As New ArrayList
    
    List.AddRange Source
    
    Assert.That List.Count, Iz.EqualTo(3)
    Assert.That List(0), Iz.EqualTo(1)
    Assert.That List(1), Iz.EqualTo(2)
    Assert.That List(2), Iz.EqualTo(3)
End Sub

Public Sub AddRange_WithICollection_AppendsToList()
    Dim Source As New ArrayList
    Source.AddRange NewLongs(1, 2, 3)
    Dim List As New ArrayList
    
    List.AddRange Source
    
    Assert.That List.Count, Iz.EqualTo(3)
    Assert.That List(0), Iz.EqualTo(1)
    Assert.That List(1), Iz.EqualTo(2)
    Assert.That List(2), Iz.EqualTo(3)
End Sub

Public Sub Set_WithExistingItem_ReplacesItem()
    Dim List As New ArrayList
    List.AddRange NewLongs(1, 2, 3)
    
    List(2) = "Hello"
    
    Assert.That List, Iz.EqualTo(Array(1&, 2&, "Hello"))
End Sub

Public Sub Clear_WhenCalled_RemovedItemsFromList()
    Dim List As New ArrayList
    List.Add 1
    
    List.Clear
    
    Assert.That List, Iz.Empty
End Sub

Public Sub CopyTo_WhenCalled_CopiesItems()
    Dim List As New ArrayList
    List.Add "Item1"
    List.Add "Item2"
    List.Add "Item3"
    Dim Target(2) As String
    
    List.CopyTo Target
    
    Assert.That Target, Iz.EqualTo(NewStrings("Item1", "Item2", "Item3"))
End Sub

Public Sub CopyTo_WithIndexOffset_CopiesItems()
    Dim List As New ArrayList
    List.Add "Item1"
    List.Add "Item2"
    List.Add "Item3"
    Dim Target(5) As String
    
    List.CopyTo Target, 2
    
    Assert.That Target, Iz.EqualTo(NewStrings(vbNullString, vbNullString, "Item1", "Item2", "Item3", vbNullString))
End Sub


Public Sub Enumerator_WhenCalled_ReturnsExpected()
    Dim List As New ArrayList
    List.AddRange NewStrings("Item1", "Item2", "Item3")
    Dim en As IEnumerator
    
    Set en = List.GetEnumerator
    
    Assert.That en, Iz.Not.Nothing
    Assert.That en.MoveNext, Iz.True
    Assert.That en.Current, Iz.EqualTo("Item1")
    Assert.That en.MoveNext, Iz.True
    Assert.That en.Current, Iz.EqualTo("Item2")
    Assert.That en.MoveNext, Iz.True
    Assert.That en.Current, Iz.EqualTo("Item3")
    Assert.That en.MoveNext, Iz.False
End Sub

Public Sub InsertRange_WhenArray_InsertsItems()
    Dim List As New ArrayList
    List.AddRange NewLongs(1, 2, 3, 4)
    
    List.Insertrange 2, NewStrings("Item1", "Item2")
    
    Assert.That List, Iz.EqualTo(Array(1&, 2&, "Item1", "Item2", 3&, 4&))
End Sub

Public Sub InsertRange_WithCollection_InsertsItems()
    Dim List As New ArrayList
    List.AddRange NewLongs(1, 2, 3, 4)
    Dim col As New Collection
    col.Add "Item1"
    col.Add "Item2"
    
    List.Insertrange 2, col
    
    Assert.That List, Iz.EqualTo(Array(1&, 2&, "Item1", "Item2", 3&, 4&))
End Sub

Public Sub Contains_WithExistingValue_ReturnsTrue()
    Dim List As New ArrayList
    List.AddRange NewLongs(1, 2, 3)
    
    Dim Actual As Boolean
    Actual = List.Contains(1&)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Contains_WithNonExistingValue_ReturnsFalse()
    Dim List As New ArrayList
    List.AddRange NewLongs(1, 2, 3)
    
    Dim Actual As Boolean
    Actual = List.Contains(4&)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub ToArray_WithEmptyList_ReturnsEmptyArray()
    Dim List As New ArrayList
    Dim l() As Long
    
    l = List.ToArray(vbLong)
    
    Assert.That l, Iz.Empty
End Sub

Public Sub LastIndexOf_WithValue_ReturnsExpected(ByVal Value As Long, ByVal Expected As Long)
    Dim List As ArrayList
    Set List = MakeFilledList
    
    Dim Actual As Long
    Actual = List.LastIndexOf(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub RemoveAt_WithValues_RemovesExpectedItem(ByVal Index As Long, ByRef Expected As Variant)
    Dim List As ArrayList
    Set List = MakeFilledList
    
    List.RemoveAt Index
    
    Assert.That List, Iz.EqualTo(Expected)
End Sub

Public Sub RemoveAt_WithInvalidIndex_Throws()
    Dim List As ArrayList
    Set List = MakeFilledList
    
    On Error Resume Next
    List.RemoveAt -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Remove_WithValue_RemovesExpected()
    Dim List As ArrayList
    Set List = MakeFilledList
    
    List.Remove 3&
    
    Assert.That List, Iz.EqualTo(NewLongs(1, 2, 4, 5, 1, 2, 3, 4, 5))
End Sub

Public Sub RemoveRange_WithValidRange_RemovesItems()
    Dim List As ArrayList
    Set List = MakeFilledList
    
    List.RemoveRange 1, 3
    
    Assert.That List, Iz.EqualTo(NewLongs(1, 5, 1, 2, 3, 4, 5))
End Sub

Public Sub RemoveRange_WithNegativeIndex_Throws()
    Dim List As ArrayList
    Set List = MakeFilledList
    
    On Error Resume Next
    List.RemoveRange -1, 3

    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub RemoveRange_WithNegativeCount_Throws()
    Dim List As ArrayList
    Set List = MakeFilledList
    
    On Error Resume Next
    List.RemoveRange 3, -1

    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub RemoveRange_WithRangeExtendingPassedEnd_Throws()
    Dim List As ArrayList
    Set List = MakeFilledList
    
    On Error Resume Next
    List.RemoveRange 9, 2

    AssertArgumentException Err
End Sub

































Private Function MakeFilledList() As ArrayList
    Set MakeFilledList = New ArrayList
    MakeFilledList.AddRange NewLongs(1, 2, 3, 4, 5, 1, 2, 3, 4, 5)
End Function

