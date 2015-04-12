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

Public Sub Adapter_WithNothing_ThrowsArgumentNullException()
    On Error GoTo errTrap
    
    ArrayList.Adapter Nothing
    
errTrap:
    AssertArgumentNullException Err, "List"
End Sub

Public Sub Count_WithEmptyList_ReturnsZero()
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(New ArrayList)
    
    Dim Actual As Long
    Actual = Adapted.Count
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub Count_WithListOfOneItem_ReturnsOne()
    Dim List As New ArrayList
    List.Add "one"
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(List)
    
    Dim Actual As Long
    Actual = Adapted.Count
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub Add_WhenCalled_AddsToList()
    Dim List As New ArrayList
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(List)
    
    Adapted.Add 1
    
    Assert.That List.Count, Iz.EqualTo(1)
    Assert.That List(0), Iz.EqualTo(1)
End Sub

Public Sub AddRange_WithCollection_AddsToList()
    Dim List As New ArrayList
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(List)
    Dim Source As New Collection
    Source.Add 1
    Source.Add 2
    Source.Add 3
    
    Adapted.AddRange Source
    
    Assert.That List, Iz.EqualTo(Source)
End Sub

Public Sub AddRange_WithArray_AddsToList()
    Dim List As New ArrayList
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(List)
    Dim Source() As Long
    Source = NewLongs(1, 2, 3)
    
    Adapted.AddRange Source
    
    Assert.That List, Iz.EqualTo(Source)
End Sub

Public Sub AddRange_WithICollection_AddsToList()
    Dim List As New ArrayList
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(List)
    Dim Source As New ArrayList
    Source.AddRange NewLongs(1, 2, 3)
    
    Adapted.AddRange Source
    
    Assert.That List, Iz.EqualTo(Source)
End Sub

Public Sub Item_WhenReading_ReturnsUnderlyingItem()
    Dim List As New ArrayList
    List.Add "one"
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(List)
    
    Dim Actual As String
    Actual = Adapted(0)
    
    Assert.That Actual, Iz.EqualTo("one")
End Sub

Public Sub Item_WhenSetting_SetsUnderlyingItem()
    Dim List As New ArrayList
    List.Add 1
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(List)
    
    Adapted(0) = 2
    
    Assert.That List(0), Iz.EqualTo(2)
End Sub

Public Sub Clear_WhenCalled_ClearsList()
    Dim List As New ArrayList
    List.Add 1
    Dim Adapted As ArrayList
    Set Adapted = ArrayList.Adapter(List)
    
    Adapted.Clear
    
    Assert.That List, Iz.Empty
End Sub

Public Sub CopyTo_WhenCalled_CopiesToOutput()
    Dim Target(2) As String
    Dim List As New ArrayList
    List.Add "Item1"
    List.Add "Item2"
    List.Add "Item3"
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)
    
    Adapter.CopyTo Target
    
    Assert.That Target, Iz.EqualTo(List)
End Sub

Public Sub CopyTo_WithIndexOffset_CopiesToOffsetOutput()
    Dim Target(5) As String
    Dim List As New ArrayList
    List.Add "Item1"
    List.Add "Item2"
    List.Add "Item3"
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)
    
    Adapter.CopyTo Target, 2
    
    Assert.That Target, Iz.EqualTo(Array("", "", "Item1", "Item2", "Item3", ""))
End Sub

Public Sub CopyTo_WithArrayIndexLessThanLowerBound_ThrowsArgumentOutOfRangeException()
    Dim Target(1 To 6) As String
    Dim List As New ArrayList
    List.Add "Item1"
    List.Add "Item2"
    List.Add "Item3"
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)
    
    On Error Resume Next
    Adapter.CopyTo Target, 0

    AssertArgumentOutOfRangeException Err, "ArrayIndex"
End Sub

Public Sub GetEnumerator_WhenCalled_ReturnsEnumerableSequence()
    Dim List As New ArrayList
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)
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
    Dim List As New ArrayList
    List.AddRange NewLongs(1, 2, 3, 4)
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)
    
    Adapter.Insertrange 0, NewStrings("Item1", "Item2", "Item3")
    
    Assert.That List, Iz.EqualTo(Array("Item1", "Item2", "Item3", 1, 2, 3, 4))
End Sub

Public Sub InsertRange_WithCollection_InsertsValues()
    Dim List As New ArrayList
    List.AddRange NewLongs(1, 2, 3, 4)
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)
    Dim Col As New Collection
    Col.Add "Item1"
    Col.Add "Item2"
    Col.Add "Item3"
    
    Adapter.Insertrange 2, Col
    
    Assert.That List, Iz.EqualTo(Array(1, 2, "Item1", "Item2", "Item3", 3, 4))
End Sub

Public Sub Contains_WithExistingValue_ReturnsTrue()
    Dim List As New ArrayList
    List.AddRange NewLongs(1, 2, 3, 4)
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)

    Dim Actual As Boolean
    Actual = Adapter.Contains(1&)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Contains_WithNonExistingValue_ReturnsFalse()
    Dim List As New ArrayList
    List.AddRange NewLongs(1, 2, 3, 4)
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)

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
    Dim List As New ArrayList
    List.AddRange Elements
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)
    
    Dim Actual As Long
    Actual = Adapter.BinarySearch(ValueToFind)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BinarySearchEx_WithSubsetValues_ReturnsExpected(ByRef Elements As Variant, ByVal ValueToFind As Integer, ByVal StartIndex As Long, ByVal Count As Long, ByVal Expected As Long)
    Dim List As New ArrayList
    List.AddRange Elements
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)
    
    Dim Actual As Long
    Actual = Adapter.BinarySearchEx(StartIndex, Count, ValueToFind)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BinarySearchEx_WithCountTooHigh_ThrowsArgumentException()
    Dim List As New ArrayList
    List.Capacity = 10
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)

    On Error Resume Next
    Adapter.BinarySearchEx 0, 11, 1
    
    AssertArgumentException Err
End Sub

Public Sub BinarySearchEx_WithNegativeStartIndex_ThrowsArgumentOutOfRangeException()
    Dim List As New ArrayList
    List.Capacity = 10
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)

    On Error Resume Next
    Adapter.BinarySearchEx -1, 1, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub BinarySearchEx_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    Dim List As New ArrayList
    List.Capacity = 10
    Dim Adapter As ArrayList
    Set Adapter = ArrayList.Adapter(List)

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
    Dim List As New ArrayList
    List.AddRange Array(1, 2, 3, 4, 1, 2, 3, 4)
    Set MakeFilledAdapter = ArrayList.Adapter(List)
End Function
