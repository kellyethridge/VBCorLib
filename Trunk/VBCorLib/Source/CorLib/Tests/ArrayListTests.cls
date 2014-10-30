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
        Case "IndexOf_WithValue_ReturnsExpected"
            Test.Use 1, 0
            Test.Use 5, 4
            Test.Use 6, -1
            
        Case "IndexOf_WithSubsetSearch_ReturnsExpected"
            Test.Use 1, 5
            Test.Use 5, 4
            Test.Use 2, -1
    
        Case "LastIndexOf_WithValue_ReturnsExpected"
            Test.Use 1, 5
            Test.Use 5, 9
            Test.Use 8, -1
            
        Case "LastIndexOf_WithSubsetSearch_ReturnsExpected"
            Test.Use 4, 3
            Test.Use 5, 4
            Test.Use 3, -1
            
        Case "RemoveAt_WithValues_RemovesExpectedItem"
            Test.Use 0, NewLongs(2, 3, 4, 5, 1, 2, 3, 4, 5)
            Test.Use 9, NewLongs(1, 2, 3, 4, 5, 1, 2, 3, 4)
            Test.Use 5, NewLongs(1, 2, 3, 4, 5, 2, 3, 4, 5)
            
        Case "Insert_WithValue_InsertsIntoList"
            Test.Use 0, "a", Array("a", 1, 2, 3, 4, 5, 1, 2, 3, 4, 5)
            Test.Use 10, "a", Array(1, 2, 3, 4, 5, 1, 2, 3, 4, 5, "a")
            Test.Use 5, "a", Array(1, 2, 3, 4, 5, "a", 1, 2, 3, 4, 5)
            
        Case "SetRange_WhenItems_SetsRangeOfItems"
            Test.Use NewStrings("Item1", "Item2")
            Test.Use Array("Item1", "Item2")
            Dim CollectionSource As New Collection
            CollectionSource.Add "Item1"
            CollectionSource.Add "Item2"
            Test.Use CollectionSource
            Dim ICollectionSource As New ArrayList
            ICollectionSource.Add "Item1"
            ICollectionSource.Add "Item2"
            Test.Use ICollectionSource
            
        Case "BinarySearch_WithValue_ReturnsExpected"
            Test.Use 1, 0
            Test.Use 7, 5
            Test.Use 4, -4
            Test.Use 0, -1
            Test.Use 8, -7
            
        Case "BinarySearch_WithRangedValue_ReturnsExpected"
            Test.Use 1, -2
            Test.Use 7, -6
            Test.Use 4, -4
            Test.Use 0, -2
            Test.Use 8, -6
    End Select
End Sub

Public Sub Constructor_WhenCalled_CreatesEmptyList()
    Dim list As ArrayList
    
    Set list = New ArrayList
    
    Assert.That list.Count, Iz.EqualTo(0)
End Sub

Public Sub Add_WithValue_IncreasesCount()
    Dim list As New ArrayList
    
    list.Add 1
    
    Assert.That list.Count, Iz.EqualTo(1)
End Sub

Public Sub Add_WithValue_AppendsToList()
    Dim list As New ArrayList
    
    list.Add 1
    
    Assert.That list.Item(0), Iz.EqualTo(1)
End Sub

Public Sub Add_WithMultipleValues_AppendsToList()
    Dim list As New ArrayList
    
    list.Add 1
    list.Add "Item2"
    
    Assert.That list(0), Iz.EqualTo(1)
    Assert.That list(1), Iz.EqualTo("Item2")
End Sub

Public Sub AddRange_WithCollection_AppendsToList()
    Dim Source As New Collection
    Source.Add 1
    Source.Add 2
    Source.Add 3
    Dim list As New ArrayList
    
    list.AddRange Source
    
    Assert.That list.Count, Iz.EqualTo(3)
    Assert.That list(0), Iz.EqualTo(1)
    Assert.That list(1), Iz.EqualTo(2)
    Assert.That list(2), Iz.EqualTo(3)
End Sub

Public Sub AddRange_WithArray_AppendsToList()
    Dim Source() As Long
    Source = NewLongs(1, 2, 3)
    Dim list As New ArrayList
    
    list.AddRange Source
    
    Assert.That list.Count, Iz.EqualTo(3)
    Assert.That list(0), Iz.EqualTo(1)
    Assert.That list(1), Iz.EqualTo(2)
    Assert.That list(2), Iz.EqualTo(3)
End Sub

Public Sub AddRange_WithICollection_AppendsToList()
    Dim Source As New ArrayList
    Source.AddRange NewLongs(1, 2, 3)
    Dim list As New ArrayList
    
    list.AddRange Source
    
    Assert.That list.Count, Iz.EqualTo(3)
    Assert.That list(0), Iz.EqualTo(1)
    Assert.That list(1), Iz.EqualTo(2)
    Assert.That list(2), Iz.EqualTo(3)
End Sub

Public Sub Set_WithExistingItem_ReplacesItem()
    Dim list As New ArrayList
    list.AddRange NewLongs(1, 2, 3)
    
    list(2) = "Hello"
    
    Assert.That list, Iz.EqualTo(Array(1&, 2&, "Hello"))
End Sub

Public Sub Clear_WhenCalled_RemovedItemsFromList()
    Dim list As New ArrayList
    list.Add 1
    
    list.Clear
    
    Assert.That list, Iz.Empty
End Sub

Public Sub CopyTo_WhenCalled_CopiesItems()
    Dim list As New ArrayList
    list.Add "Item1"
    list.Add "Item2"
    list.Add "Item3"
    Dim Target(2) As String
    
    list.CopyTo Target
    
    Assert.That Target, Iz.EqualTo(NewStrings("Item1", "Item2", "Item3"))
End Sub

Public Sub CopyTo_WithIndexOffset_CopiesItems()
    Dim list As New ArrayList
    list.Add "Item1"
    list.Add "Item2"
    list.Add "Item3"
    Dim Target(5) As String
    
    list.CopyTo Target, 2
    
    Assert.That Target, Iz.EqualTo(NewStrings(vbNullString, vbNullString, "Item1", "Item2", "Item3", vbNullString))
End Sub


Public Sub GetEnumerator_WhenCalled_ReturnsEnumeration()
    Dim list As New ArrayList
    list.AddRange NewStrings("Item1", "Item2", "Item3")
    Dim en As IEnumerator
    
    Set en = list.GetEnumerator
    
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
    Dim list As New ArrayList
    list.AddRange NewLongs(1, 2, 3, 4)
    
    list.Insertrange 2, NewStrings("Item1", "Item2")
    
    Assert.That list, Iz.EqualTo(Array(1&, 2&, "Item1", "Item2", 3&, 4&))
End Sub

Public Sub InsertRange_WithCollection_InsertsItems()
    Dim list As New ArrayList
    list.AddRange NewLongs(1, 2, 3, 4)
    Dim col As New Collection
    col.Add "Item1"
    col.Add "Item2"
    
    list.Insertrange 2, col
    
    Assert.That list, Iz.EqualTo(Array(1&, 2&, "Item1", "Item2", 3&, 4&))
End Sub

Public Sub Contains_WithExistingValue_ReturnsTrue()
    Dim list As New ArrayList
    list.AddRange NewLongs(1, 2, 3)
    
    Dim Actual As Boolean
    Actual = list.Contains(1&)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Contains_WithNonExistingValue_ReturnsFalse()
    Dim list As New ArrayList
    list.AddRange NewLongs(1, 2, 3)
    
    Dim Actual As Boolean
    Actual = list.Contains(4&)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub ToArray_WithEmptyList_ReturnsEmptyArray()
    Dim list As New ArrayList
    Dim l() As Long
    
    l = list.ToArray(vbLong)
    
    Assert.That l, Iz.Empty
End Sub

Public Sub IndexOf_WithValue_ReturnsExpected(ByVal Value As Long, ByVal Expected As Long)
    Dim list As ArrayList
    Set list = MakeFilledList
    
    Dim Actual As Long
    Actual = list.IndexOf(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IndexOf_WithSubsetSearch_ReturnsExpected(ByVal Value As Long, ByVal Expected As Long)
    Dim list As ArrayList
    Set list = MakeFilledList
    
    Dim Actual As Long
    Actual = list.IndexOf(Value, 2, 4)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IndexOf_WithNegativeStartIndex_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.IndexOf 1, -1, 1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub IndexOf_WithStartIndexEqualToCount_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.IndexOf 1, 10
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub IndexOf_WithNegativeCount_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.IndexOf 1, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub IndexOf_WithInvalidRange_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.IndexOf 1, 0, 11
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub IndexOf_WithStartIndex_ReturnsExpected()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    Dim Actual As Long
    Actual = list.IndexOf(1&, 1)
    
    Assert.That Actual, Iz.EqualTo(5)
End Sub

Public Sub LastIndexOf_WithValue_ReturnsExpected(ByVal Value As Long, ByVal Expected As Long)
    Dim list As ArrayList
    Set list = MakeFilledList
    
    Dim Actual As Long
    Actual = list.LastIndexOf(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub LastIndexOf_WithSubsetSearch_ReturnsExpected(ByVal Value As Long, ByVal Expected As Long)
    Dim list As ArrayList
    Set list = MakeFilledList
    
    Dim Actual As Long
    Actual = list.LastIndexOf(Value, 6, 4)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub LastIndexOf_WithNegativeStartIndex_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.LastIndexOf 1, -1, 1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub LastIndexOf_WithStartIndexEqualToCount_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.LastIndexOf 1, 10, 1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub LastIndexOf_WithNegativeCount_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.LastIndexOf 1, 2, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub LastIndexOf_WithInvalidRange_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.LastIndexOf 1, 2, 4
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub LastIndexOf_WithStartIndex_ReturnsExpected()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    Dim Actual As Long
    Actual = list.LastIndexOf(5&, 8)
    
    Assert.That Actual, Iz.EqualTo(4)
End Sub

Public Sub RemoveAt_WithValues_RemovesExpectedItem(ByVal Index As Long, ByRef Expected As Variant)
    Dim list As ArrayList
    Set list = MakeFilledList
    
    list.RemoveAt Index
    
    Assert.That list, Iz.EqualTo(Expected)
End Sub

Public Sub RemoveAt_WithInvalidIndex_ThrowsArgumentOutOfRangeException()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.RemoveAt -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Remove_WithValue_RemovesExpected()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    list.Remove 3&
    
    Assert.That list, Iz.EqualTo(NewLongs(1, 2, 4, 5, 1, 2, 3, 4, 5))
End Sub

Public Sub RemoveRange_WithValidRange_RemovesItems()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    list.RemoveRange 1, 3
    
    Assert.That list, Iz.EqualTo(NewLongs(1, 5, 1, 2, 3, 4, 5))
End Sub

Public Sub RemoveRange_WithNegativeIndex_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.RemoveRange -1, 3

    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub RemoveRange_WithNegativeCount_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.RemoveRange 3, -1

    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub RemoveRange_WithRangeExtendingPassedEnd_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.RemoveRange 9, 2

    AssertArgumentException Err
End Sub

Public Sub Reverse_WhenCalled_ReversesItems()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    list.Reverse
    
    Assert.That list, Iz.EqualTo(NewLongs(5, 4, 3, 2, 1, 5, 4, 3, 2, 1))
End Sub

Public Sub Reverse_WithRangeSpecified_ReversesItemsInRange()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    list.Reverse 2, 3
    
    Assert.That list, Iz.EqualTo(NewLongs(1, 2, 5, 4, 3, 1, 2, 3, 4, 5))
End Sub

Public Sub Reverse_WithNegativeIndex_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.Reverse -1, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Reverse_WithNegativeCount_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.Reverse 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Reverse_WithIndexAndCountExceedLength_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.Reverse 5, 6
    
    AssertArgumentException Err
End Sub

Public Sub ToArray_WhenCalled_ReturnsArrayOfItems()
    Dim list As ArrayList
    Set list = MakeFilledList
    Dim Arr() As Long
    
    Arr = list.ToArray(vbLong)
    
    Assert.That Arr, Iz.EqualTo(NewLongs(1, 2, 3, 4, 5, 1, 2, 3, 4, 5))
End Sub

Public Sub SetRange_WhenItems_SetsRangeOfItems(ByRef Items As Variant)
    Dim list As ArrayList
    Set list = MakeFilledList
    
    list.SetRange 2, Items
    
    Assert.That list, Iz.EqualTo(Array(1, 2, "Item1", "Item2", 5, 1, 2, 3, 4, 5))
End Sub

Public Sub SetRange_WithNullCollection_Throws()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.SetRange 0, Nothing
    
    AssertArgumentNullException Err, "c"
End Sub

Public Sub GetEnumerator_WithSubset_EnumeratesSubset()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    Dim Actual As Object
    Set Actual = list.GetEnumerator(2, 4)
    
    Assert.That Actual, Iz.EqualTo(Array(3, 4, 5, 1))
End Sub

Public Sub Insert_WithValue_InsertsIntoList(ByVal Index As Long, ByVal Value As Variant, ByRef Expected As Variant)
    Dim list As ArrayList
    Set list = MakeFilledList
    
    list.Insert Index, Value
    
    Assert.That list, Iz.EqualTo(Expected)
End Sub

Public Sub Insert_WithEmptyList_InsertsValue()
    Dim list As New ArrayList
    
    list.Insert 0, "a"
    
    Assert.That list, Iz.EqualTo(Array("a"))
End Sub

Public Sub Insert_WithNegativeIndex_ThrowsArgumentOutOfRangeException()
    Dim list As New ArrayList
    
    On Error Resume Next
    list.Insert -1, "a"
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Insert_WithIndexLargerThanListSize_ThrowsArgumentOutOfRangeException()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    On Error Resume Next
    list.Insert 11, "a"
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub TestSerialization()
    Dim SourceList As ArrayList
    Set SourceList = MakeFilledList
    Dim SourceBag As New PropertyBag
    
    SourceBag.WriteProperty "List", SourceList
    
    Dim TargetList As ArrayList
    Dim TargetBag As New PropertyBag
    
    TargetBag.Contents = SourceBag.Contents
    Set TargetList = TargetBag.ReadProperty("List", Nothing)
    
    Assert.That TargetList, Iz.EqualTo(Array(1, 2, 3, 4, 5, 1, 2, 3, 4, 5))
End Sub

Public Sub BinarySearch_WithValue_ReturnsExpected(ByVal Value As Long, ByVal Expected As Long)
    Dim list As ArrayList
    Set list = MakeOrderedList
    
    Dim Actual As Long
    Actual = list.BinarySearch(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BinarySearchEx_WithRangedValue_ReturnsExpected(ByVal Value As Long, ByVal Expected As Long)
    Dim list As ArrayList
    Set list = MakeOrderedList
    
    Dim Actual As Long
    Actual = list.BinarySearchEx(1, 4, Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BinarySearchEx_WithNegativeIndex_Throws()
    Dim list As ArrayList
    Set list = MakeOrderedList
    
    On Error Resume Next
    list.BinarySearchEx -1, 1, 4
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub BinarySearchEx_WithNegativeCount_Throws()
    Dim list As ArrayList
    Set list = MakeOrderedList
    
    On Error Resume Next
    list.BinarySearchEx 0, -1, 4
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub BinarySearchEx_WithInvalidRange_Throws()
    Dim list As ArrayList
    Set list = MakeOrderedList
    
    On Error Resume Next
    list.BinarySearchEx 1, 6, 4
    
    AssertArgumentException Err
End Sub

Public Sub Sort_WhenCalled_OrdersItems()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    list.Sort
    
    Assert.That list, Iz.Ordered
End Sub

Public Sub Sort_WithSortRange_OrdersItemsInRange()
    Dim list As ArrayList
    Set list = MakeFilledList
    
    list.Sort 2, 5
    
    Assert.That list, Iz.EqualTo(NewLongs(1, 2, 1, 2, 3, 4, 5, 3, 4, 5))
End Sub














Private Function MakeFilledList() As ArrayList
    Set MakeFilledList = New ArrayList
    MakeFilledList.AddRange NewLongs(1, 2, 3, 4, 5, 1, 2, 3, 4, 5)
End Function

Private Function MakeOrderedList() As ArrayList
    Set MakeOrderedList = New ArrayList
    MakeOrderedList.AddRange NewLongs(1, 2, 3, 5, 6, 7)
End Function
