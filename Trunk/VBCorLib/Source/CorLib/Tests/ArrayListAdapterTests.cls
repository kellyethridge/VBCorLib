VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayListAdapterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ArrayListAdapterTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "BinarySearch_WithValues_ReturnsExpected"
            Test.Use Array(1, 2, 3, 4, 5), 4, 3
            Test.Use Array(1, 2, 3, 4, 5), 1, 0
            Test.Use Array(1, 2, 3, 4, 5), 5, 4
            Test.Use Array(1, 2, 3, 4, 5), 6, -6
            Test.Use Array(1, 2, 4, 5), 3, -3
            Test.Use Array(1, 2, 3, 4, 5), 0, -1
            
        Case "BinarySearch_WithSubsetValues_ReturnsExpected"
            Test.Use Array(1, 2, 3, 4, 5, 6, 7, 8), 4, 2, 6, 3
            Test.Use Array(1, 2, 3, 4, 5, 6, 7, 8), 4, 4, 4, -5
            Test.Use Array(1, 2, 3, 4, 5, 6, 7, 8), 7, 1, 5, -7
            Test.Use Array(1, 2, 3, 5, 6, 7, 8), 4, 2, 5, -4
    End Select
End Sub

Public Sub Count_WithEmptyList_ReturnsZero()
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(New ArrayList)
    
    Dim Actual As Long
    Actual = Adapted.Count
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub Count_WithListOfOneItem_ReturnsOne()
    Dim list As New ArrayList
    list.Add "one"
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(list)
    
    Dim Actual As Long
    Actual = Adapted.Count
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub Add_WhenCalled_AddsToList()
    Dim list As New ArrayList
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(list)
    
    Adapted.Add 1
    
    Assert.That list.Count, Iz.EqualTo(1)
    Assert.That list(0), Iz.EqualTo(1)
End Sub

Public Sub AddRange_WithCollection_AddsToList()
    Dim list As New ArrayList
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(list)
    Dim Source As New Collection
    Source.Add 1
    Source.Add 2
    Source.Add 3
    
    Adapted.AddRange Source
    
    Assert.That list, Iz.EqualTo(Source)
End Sub

Public Sub AddRange_WithArray_AddsToList()
    Dim list As New ArrayList
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(list)
    Dim Source() As Long
    Source = NewLongs(1, 2, 3)
    
    Adapted.AddRange Source
    
    Assert.That list, Iz.EqualTo(Source)
End Sub

Public Sub AddRange_WithICollection_AddsToList()
    Dim list As New ArrayList
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(list)
    Dim Source As New ArrayList
    Source.AddRange NewLongs(1, 2, 3)
    
    Adapted.AddRange Source
    
    Assert.That list, Iz.EqualTo(Source)
End Sub

Public Sub Item_WhenReading_ReturnsUnderlyingItem()
    Dim list As New ArrayList
    list.Add "one"
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(list)
    
    Dim Actual As String
    Actual = Adapted(0)
    
    Assert.That Actual, Iz.EqualTo("one")
End Sub

Public Sub Item_WhenSetting_SetsUnderlyingItem()
    Dim list As New ArrayList
    list.Add 1
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(list)
    
    Adapted(0) = 2
    
    Assert.That list(0), Iz.EqualTo(2)
End Sub

Public Sub Clear_WhenCalled_ClearsList()
    Dim list As New ArrayList
    list.Add 1
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(list)
    
    Adapted.Clear
    
    Assert.That list, Iz.Empty
End Sub

Public Sub CopyTo_WhenCalled_CopiesToOutput()
    Dim Target(2) As String
    Dim list As New ArrayList
    list.Add "Item1"
    list.Add "Item2"
    list.Add "Item3"
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)
    
    Adapter.CopyTo Target
    
    Assert.That Target, Iz.EqualTo(list)
End Sub

Public Sub CopyTo_WithIndexOffset_CopiesToOffsetOutput()
    Dim Target(5) As String
    Dim list As New ArrayList
    list.Add "Item1"
    list.Add "Item2"
    list.Add "Item3"
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)
    
    Adapter.CopyTo Target, 2
    
    Assert.That Target, Iz.EqualTo(Array("", "", "Item1", "Item2", "Item3", ""))
End Sub

Public Sub CopyTo_WithArrayIndexLessThanLowerBound_ThrowsArgumentOutOfRangeException()
    Dim Target(1 To 6) As String
    Dim list As New ArrayList
    list.Add "Item1"
    list.Add "Item2"
    list.Add "Item3"
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)
    
    On Error Resume Next
    Adapter.CopyTo Target, 0

    AssertArgumentOutOfRangeException Err, "ArrayIndex"
End Sub

Public Sub GetEnumerator_WhenCalled_ReturnsEnumerableSequence()
    Dim list As New ArrayList
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)
    Adapter.AddRange Array("Item1", "Item2", "Item3")
    
    Dim Actual As IEnumerator
    Set Actual = Adapter.GetEnumerator
    
    Assert.That Actual.MoveNext, Iz.True
    Assert.That Actual.Current, Iz.EqualTo("Item1")
    Assert.That Actual.MoveNext, Iz.True
    Assert.That Actual.Current, Iz.EqualTo("Item2")
    Assert.That Actual.MoveNext, Iz.True
    Assert.That Actual.Current, Iz.EqualTo("Item3")
    Assert.That Actual.MoveNext, Iz.False
End Sub

Public Sub GetEnumerater_WithSubset_REturnsEnumerableSubset()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter
    
    Dim Actual As IEnumerator
    Set Actual = Adapter.GetEnumerator(2, 3)
    
    Assert.That Actual.MoveNext, Iz.True
    Assert.That Actual.Current, Iz.EqualTo(3)
    Assert.That Actual.MoveNext, Iz.True
    Assert.That Actual.Current, Iz.EqualTo(4)
    Assert.That Actual.MoveNext, Iz.True
    Assert.That Actual.Current, Iz.EqualTo(1)
    Assert.That Actual.MoveNext, Iz.False
End Sub

Public Sub Insert_WithArray_InsertsValues()
    Dim list As New ArrayList
    list.AddRange NewLongs(1, 2, 3, 4)
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)
    
    Adapter.Insertrange 0, NewStrings("Item1", "Item2", "Item3")
    
    Assert.That list, Iz.EqualTo(Array("Item1", "Item2", "Item3", 1, 2, 3, 4))
End Sub

Public Sub InsertRange_WithCollection_InsertsValues()
    Dim list As New ArrayList
    list.AddRange NewLongs(1, 2, 3, 4)
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)
    Dim col As New Collection
    col.Add "Item1"
    col.Add "Item2"
    col.Add "Item3"
    
    Adapter.Insertrange 2, col
    
    Assert.That list, Iz.EqualTo(Array(1, 2, "Item1", "Item2", "Item3", 3, 4))
End Sub

Public Sub Contains_WithExistingValue_ReturnsTrue()
    Dim list As New ArrayList
    list.AddRange NewLongs(1, 2, 3, 4)
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)

    Dim Actual As Boolean
    Actual = Adapter.Contains(1&)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Contains_WithNonExistingValue_ReturnsFalse()
    Dim list As New ArrayList
    list.AddRange NewLongs(1, 2, 3, 4)
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)

    Dim Actual As Boolean
    Actual = Adapter.Contains(5)

    Assert.That Actual, Iz.False
End Sub

Public Sub ToArray_WithEmptyList_ReturnsEmptyArray()
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(New ArrayList)
    
    Dim Actual() As Long
    Actual = Adapter.ToArray(vbLong)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub ToArray_WithElements_ReturnsArrayOfElements()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter
    
    Dim Actual() As Long
    Actual = Adapter.ToArray(vbLong)
    
    Assert.That Actual, Iz.EqualTo(Array(1, 2, 3, 4, 1, 2, 3, 4))
End Sub

Public Sub SetRange_WhenCalled_ReplacesElements()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter
    
    Adapter.SetRange 2, Array(5, 6, 7, 8)
    
    Assert.That Adapter, Iz.EqualTo(Array(1, 2, 5, 6, 7, 8, 3, 4))
End Sub

Public Sub LastIndexOf_WhenCalled_ReturnsExpected()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter
    
    Dim Actual As Long
    Actual = Adapter.LastIndexOf(2)
    
    Assert.That Actual, Iz.EqualTo(5)
End Sub

Public Sub RemoveAt_WhenCalled_RemovesItem()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter
    
    Adapter.RemoveAt 1
    
    Assert.That Adapter, Iz.EqualTo(Array(1, 3, 4, 1, 2, 3, 4))
End Sub

Public Sub Remove_WhenCalled_RemovesExpected()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter

    Adapter.Remove 3
    
    Assert.That Adapter, Iz.EqualTo(Array(1, 2, 4, 1, 2, 3, 4))
End Sub

Public Sub RemoveRange_WhenCalled_RemovesExpected()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter

    Adapter.RemoveRange 1, 2
    
    Assert.That Adapter, Iz.EqualTo(Array(1, 4, 1, 2, 3, 4))
End Sub

Public Sub Reverse_WhenCalled_ReversedList()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter

    Adapter.Reverse
    
    Assert.That Adapter, Iz.EqualTo(Array(4, 3, 2, 1, 4, 3, 2, 1))
End Sub

Public Sub Reverse_WithSubset_ReversesSubset()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter

    Adapter.Reverse 2, 3
    
    Assert.That Adapter, Iz.EqualTo(Array(1, 2, 1, 4, 3, 2, 3, 4))
End Sub

Public Sub IndexOf_WhenCalled_ReturnsExpected()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter

    Dim Actual As Long
    Actual = Adapter.IndexOf(4)
    
    Assert.That Actual, Iz.EqualTo(3)
End Sub

Public Sub IndexOf_WhenNonExistingValue_ReturnsMinusOne()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter

    Dim Actual As Long
    Actual = Adapter.IndexOf(5)
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub IndexOf_WithIndexOffset_ReturnsExpected()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter

    Dim Actual As Long
    Actual = Adapter.IndexOf(1, 2, 4)
    
    Assert.That Actual, Iz.EqualTo(4)
End Sub

Public Sub Insert_WhenCalledWithFirstIndex_InsertsAtStart()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter

    Adapter.Insert 0, 5
    
    Assert.That Adapter, Iz.EqualTo(Array(5, 1, 2, 3, 4, 1, 2, 3, 4))
End Sub

Public Sub Insert_WhenCalledWithLastIndex_InsertsAtEnd()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter

    Adapter.Insert 8, 5
    
    Assert.That Adapter, Iz.EqualTo(Array(1, 2, 3, 4, 1, 2, 3, 4, 5))
End Sub

Public Sub BinarySearch_WithValues_ReturnsExpected(ByRef Elements As Variant, ByRef ValueToFind As Variant, ByVal Expected As Long)
    Dim list As New ArrayList
    list.AddRange Elements
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)
    
    Dim Actual As Long
    Actual = Adapter.BinarySearch(ValueToFind)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BinarySearchEx_WithSubsetValues_ReturnsExpected(ByRef Elements As Variant, ByVal ValueToFind As Integer, ByVal StartIndex As Long, ByVal Count As Long, ByVal Expected As Long)
    Dim list As New ArrayList
    list.AddRange Elements
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)
    
    Dim Actual As Long
    Actual = Adapter.BinarySearchEx(StartIndex, Count, ValueToFind)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BinarySearchEx_WithCountTooHigh_ThrowsArgumentException()
    Dim list As New ArrayList
    list.Capacity = 10
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)

    On Error Resume Next
    Adapter.BinarySearchEx 0, 11, 1
    
    AssertArgumentException Err
End Sub

Public Sub BinarySearchEx_WithNegativeStartIndex_ThrowsArgumentOutOfRangeException()
    Dim list As New ArrayList
    list.Capacity = 10
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)

    On Error Resume Next
    Adapter.BinarySearchEx -1, 1, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub BinarySearchEx_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    Dim list As New ArrayList
    list.Capacity = 10
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(list)

    On Error Resume Next
    Adapter.BinarySearchEx 0, -1, 1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Sort_WhenCalled_SortsElements()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter
    
    Adapter.Sort
    
    Assert.That Adapter, Iz.Ordered
End Sub

Public Sub Sort_WithSubset_SortsSubset()
    Dim Adapter As ArrayList
    Set Adapter = MakeFilledAdapter
    
    Adapter.Sort 2, 4
    
    Assert.That Adapter, Iz.EqualTo(Array(1, 2, 1, 2, 3, 4, 3, 4))
End Sub

Private Function MakeFilledAdapter() As ArrayList
    Dim list As New ArrayList
    list.AddRange Array(1, 2, 3, 4, 1, 2, 3, 4)
    Set MakeFilledAdapter = ArrayList.Adapter(list)
End Function
