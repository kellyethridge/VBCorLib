VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortedListTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' SortedListTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Contains_WithNumericKeys_ReturnsExpected"
            Test.Use 1, True
            Test.Use 2, False
            Test.Use 3, True
            Test.Use 5, True
            Test.Use 6, False
        
        Case "Contains_WithObjectKeys_ReturnsExpected"
            Test.Use NewInt32(1), True
            Test.Use NewInt32(2), False
            Test.Use NewInt32(3), True
            
        Case "ContainsValue_WithValue_ReturnsExpected"
            Test.Use "hi", True
            Test.Use "bye", False
            Test.Use 3, True
            Test.Use 4, False
            Test.Use NewInt32(7), True
            Test.Use NewInt32(8), False
            
        Case "GetKey_WithValue_ReturnsExpected"
            Test.Use 0, 1
            Test.Use 1, 2
            Test.Use 2, 3
            
        Case "GetKey_WithInvalidIndex_ThrowsArgumentOutOfRangeException"
            Test.Use -1
            Test.Use 3
            
        Case "GetByIndex_WithValue_ReturnsExpected"
            Test.Use 0, "hi"
            Test.Use 1, 3
            Test.Use 2, NewInt32(7)
        
        Case "GetByIndex_WithInvalidIndex_ThrowsArgumentOutOfRangeException"
            Test.Use -1
            Test.Use 3
            
        Case "SetByIndex_WithInvalidIndex_ThrowsArgumentOutOfRangeException"
            Test.Use -1
            Test.Use 3
        
        Case "RemoveAt_WithInvalidIndex_ThrowsArgumentOutOfRangeException"
            Test.Use -1
            Test.Use 3
            
        Case "CopyTo_WithIndex_CopiesToArray"
            Test.Use 0, Array(Cor.NewDictionaryEntry(1, "a"), Cor.NewDictionaryEntry(2, "b"), Cor.NewDictionaryEntry(3, "c"), Nothing)
            Test.Use 1, Array(Nothing, Cor.NewDictionaryEntry(1, "a"), Cor.NewDictionaryEntry(2, "b"), Cor.NewDictionaryEntry(3, "c"))
    
        Case "IndexOfKey_WithValue_ReturnsExpected"
            Test.Use 1, 0
            Test.Use 2, 1
            Test.Use 3, -1
        
        Case "IndexOfValue_WithValue_ReturnsExpected"
            Test.Use "a", 0
            Test.Use "b", 1
            Test.Use "c", -1
            
    End Select
End Sub

Public Sub Count_WhenEmpty_ReturnsZero()
    Dim List As New SortedList
    
    Dim Actual As Long
    Actual = List.Count
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub Add_WhenCalled_AddsToList()
    Dim List As New SortedList
    
    List.Add 1, "hi"
    
    Assert.That List.Count, Iz.EqualTo(1)
End Sub

Public Sub Add_WithExistingKey_ThrowsArgumentException()
    Dim List As New SortedList
    
    On Error Resume Next
    List.Add 1, "hi"
    List.Add 1, "bye"
    
    AssertArgumentException Err
End Sub

Public Sub Add_WithDifferentKeyDataTypes_ThrowsInvalidOperation()
    On Error Resume Next
    Dim List As New SortedList
    
    List.Add 1, "a"
    List.Add 2&, "b"
    
    AssertInvalidOperationException Err
End Sub

Public Sub Add_WithCapacitySetToZero_AddsItem()
    Dim List As New SortedList
    List.Capacity = 0
    
    List.Add 1, "a"
    
    Assert.That List(1), Iz.EqualTo("a")
    Assert.That List.Capacity, Iz.EqualTo(16)
End Sub

Public Sub Item_WithKey_ReturnsValue()
    Dim List As New SortedList
    List.Add "1", "hi"
    List.Add "2", "Hello"
    List.Add "key", "Me"
    
    Assert.That List("key"), Iz.EqualTo("Me")
    Assert.That List("2"), Iz.EqualTo("Hello")
    Assert.That List("1"), Iz.EqualTo("hi")
End Sub

Public Sub Item_WhenSet_ReplacesExistingValue()
    Dim List As New SortedList
    List.Add "1", "hi"
    
    List("1") = "bye"
    
    Assert.That List("1"), Iz.EqualTo("bye")
End Sub

Public Sub Item_WhenSetToObject_ReplacesExistingValue()
    Dim List As New SortedList
    List.Add "1", "hi"
    
    Set List("1") = Me
    
    Assert.That List("1"), Iz.SameAs(Me)
End Sub

Public Sub Item_WhenSet_AddsNonExisingValue()
    Dim List As New SortedList
    
    List("1") = "hi"
    
    Assert.That List("1"), Iz.EqualTo("hi")
End Sub

Public Sub Item_WhenSetWithObject_AddsNonExistingValue()
    Dim List As New SortedList
    
    Set List("1") = Me
    
    Assert.That List("1"), Iz.SameAs(Me)
End Sub

Public Sub GetByIndex_WhenCalled_ReturnsInSortedOrder()
    Dim List As New SortedList
    List.Add 3, "Me"
    List.Add 1, "hi"
    List.Add 2, "Hello"
    
    Assert.That List.GetByIndex(0), Iz.EqualTo("hi")
    Assert.That List.GetByIndex(1), Iz.EqualTo("Hello")
    Assert.That List.GetByIndex(2), Iz.EqualTo("Me")
End Sub

Public Sub Contains_WithNumericKeys_ReturnsExpected(ByRef Key As Variant, ByVal Expected As Boolean)
    Dim List As New SortedList
    List.Add 1, "hi"
    List.Add 3, "bye"
    List.Add 5, "5"
    
    Dim Actual As Boolean
    Actual = List.Contains(Key)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Contains_WithObjectKeys_ReturnsExpected(ByRef Key As Variant, ByVal Expected As Boolean)
    Dim List As New SortedList
    List.Add NewInt32(1), "hi"
    List.Add NewInt32(3), "bye"
    List.Add NewInt32(5), "5"
    
    Dim Actual As Boolean
    Actual = List.Contains(Key)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ContainsValue_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal Expected As Boolean)
    Dim List As New SortedList
    List.Add 1, "hi"
    List.Add 2, 3
    List.Add 3, NewInt32(7)
    
    Dim Actual As Boolean
    Actual = List.ContainsValue(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetKey_WithValue_ReturnsExpected(ByVal Index As Long, ByVal Expected As Long)
    Dim List As New SortedList
    List.Add 3, NewInt32(7)
    List.Add 2, 3
    List.Add 1, "hi"
    
    Dim Actual As Long
    Actual = List.GetKey(Index)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetKey_WithInvalidIndex_ThrowsArgumentOutOfRangeException(ByVal InvalidIndex As Long)
    On Error Resume Next
    Dim List As New SortedList
    List.Add 3, NewInt32(7)
    List.Add 2, 3
    List.Add 1, "hi"

    List.GetKey InvalidIndex
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetByIndex_WithValue_ReturnsExpected(ByVal Index As Long, ByRef Expected As Variant)
    Dim List As New SortedList
    List.Add 3, NewInt32(7)
    List.Add 2, 3
    List.Add 1, "hi"
    
    Dim Actual As Variant
    MoveVariant Actual, List.GetByIndex(Index)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub GetByIndex_WithInvalidIndex_ThrowsArgumentOutOfRangeException(ByVal InvalidIndex As Long)
    On Error Resume Next
    Dim List As New SortedList
    List.Add 3, NewInt32(7)
    List.Add 2, 3
    List.Add 1, "hi"
    
    List.GetByIndex InvalidIndex
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub SetByIndex_WhenCalled_ReplacesValue()
    Dim List As New SortedList
    List.Add 3, NewInt32(7)
    List.Add 2, 3
    List.Add 1, "hi"

    List.SetByIndex 2, 7
    
    Assert.That List.GetByIndex(2), Iz.EqualTo(7)
End Sub

Public Sub SetByIndex_WithInvalidIndex_ThrowsArgumentOutOfRangeException(ByVal InvalidIndex As Long)
    On Error Resume Next
    Dim List As New SortedList
    List.Add 3, NewInt32(7)
    List.Add 2, 3
    List.Add 1, "hi"
    
    List.SetByIndex InvalidIndex, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Remove_WithExistingKey_RemovesItem()
    Dim List As New SortedList
    List.Add 3, NewInt32(7)
    List.Add 2, 3
    List.Add 1, "hi"

    List.Remove 1
    
    Assert.That List.Count, Iz.EqualTo(2)
    Assert.That List.GetKey(1), Iz.EqualTo(3)
    Assert.That List.GetByIndex(1), Equals(NewInt32(7))
End Sub

Public Sub Remove_WithNonExistingKey_ReturnsEmpty()
    Dim List As New SortedList
    List.Add 3, NewInt32(7)
    List.Add 2, 3
    List.Add 1, "hi"

    Dim Actual As Variant
    Actual = List(5)
    
    Assert.That Actual, Iz.vbEmpty
End Sub

Public Sub RemoveAt_WithValue_RemovesItem()
    Dim List As New SortedList
    List.Add 3, NewInt32(7)
    List.Add 2, 3
    List.Add 1, "hi"

    List.RemoveAt 1
    
    Assert.That List.Count, Iz.EqualTo(2)
    Assert.That List.GetKey(1), Iz.EqualTo(3)
    Assert.That List.GetByIndex(1), Equals(NewInt32(7))
End Sub

Public Sub RemoveAt_WithInvalidIndex_ThrowsArgumentOutOfRangeException(ByVal InvalidIndex As Long)
    On Error Resume Next
    Dim List As New SortedList
    List.Add 3, NewInt32(7)
    List.Add 2, 3
    List.Add 1, "hi"

    List.RemoveAt InvalidIndex
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub CopyTo_WithIndex_CopiesToArray(ByVal Index As Long, ByRef Expected As Variant)
    Dim List As New SortedList
    List.Add 1, "a"
    List.Add 2, "b"
    List.Add 3, "c"
    
    Dim Actual(0 To 3) As DictionaryEntry
    List.CopyTo Actual, Index
        
    Assert.That Actual, Iz.EqualTo(Expected).Using(New CorObjectComparer)
End Sub

Public Sub CopyTo_WithNullArray_ThrowsArgumentNullException()
    Dim List As New SortedList
    Dim NullArray() As Long
    
    On Error Resume Next
    List.CopyTo NullArray, 0
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub CopyTo_WithMultiDimArray_ThrowsArgumentException()
    On Error Resume Next
    Dim List As New SortedList
    Dim MultiDim(1, 1) As Variant
    
    List.CopyTo MultiDim, 0
    
    AssertArgumentException Err, "Arr"
End Sub

Public Sub CopyTo_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    Dim List As New SortedList
    Dim Target(1 To 1) As Variant
     
    List.CopyTo Target, 0
     
    AssertArgumentOutOfRangeException Err, "ArrayIndex"
End Sub

Public Sub CopyTo_WithIndexAndListCountGreaterThanUBound_ThrowsArgumentException()
    On Error Resume Next
    Dim List As New SortedList
    List.Add 1, "a"
    List.Add 2, "b"
    Dim Target(0 To 1) As Variant
    
    List.CopyTo Target, 1
    
    AssertArgumentException Err
End Sub

Public Sub CopyTo_WithIncompatibleDestinationArray_ThrowsInvalidCastException()
    Dim Dst() As Long
    Dim List As New SortedList
    
    ReDim Dst(0)
    List.Add 1, "a"
    
    On Error Resume Next
    List.CopyTo Dst, 0
    
    AssertInvalidCastException Err
End Sub

Public Sub IndexOfKey_WithValue_ReturnsExpected(ByRef Key As Variant, ByVal Expected As Long)
    Dim List As New SortedList
    List.Add 1, "a"
    List.Add 2, "b"
    
    Dim Actual As Long
    Actual = List.IndexOfKey(Key)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IndexOfValue_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal Expected As Long)
    Dim List As New SortedList
    List.Add 1, "a"
    List.Add 2, "b"
    
    Dim Actual As Long
    Actual = List.IndexOfValue(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TrimToSize_WhenCalled_SetsCapacityToCount()
    Dim List As New SortedList
    List.Capacity = 100
    List.Add 1, "a"
    List.Add 2, "b"
    
    List.TrimToSize
    
    Assert.That List.Capacity, Iz.EqualTo(2)
    Assert.That List.GetByIndex(0), Iz.EqualTo("a")
    Assert.That List.GetByIndex(1), Iz.EqualTo("b")
End Sub

Public Sub Enumeration_WhenCalled_EnumeratesList()
    Dim List As New SortedList
    List.Add 1, "a"
    List.Add 2, "b"
    List.Add 3, "c"

    Dim Expected() As Variant
    Expected = Array(Cor.NewDictionaryEntry(1, "a"), Cor.NewDictionaryEntry(2, "b"), Cor.NewDictionaryEntry(3, "c"))
    Assert.That List, Iz.EqualTo(Expected).Using(New CorObjectComparer)
End Sub

Public Sub Clone_WhenCalled_ReturnsCopyOfList()
    Dim List As New SortedList
    List.Add 1, "a"
    List.Add 2, "b"
    List.Add 3, "c"
    
    Dim Actual As SortedList
    Set Actual = List.Clone
    
    Dim Expected() As Variant
    Expected = Array(Cor.NewDictionaryEntry(1, "a"), Cor.NewDictionaryEntry(2, "b"), Cor.NewDictionaryEntry(3, "c"))
    Assert.That Actual, Iz.Not.SameAs(List)
    Assert.That Actual, Iz.EqualTo(Expected).Using(New CorObjectComparer)
End Sub

Public Sub Serialization_WhenCalled_SerializesList()
    Dim List As New SortedList
    List.Add 1, "a"
    List.Add 2, "b"
    List.Add 3, "c"
    
    Dim Actual As SortedList
    Set Actual = Persist(List)
    
    Dim Expected() As Variant
    Expected = Array(Cor.NewDictionaryEntry(1, "a"), Cor.NewDictionaryEntry(2, "b"), Cor.NewDictionaryEntry(3, "c"))
    Assert.That Actual, Iz.EqualTo(Expected).Using(New CorObjectComparer)
End Sub

Public Sub GetKeyList_WhenCalled_ReturnsListOfKeys()
    Dim List As New SortedList
    List.Add 1, "a"
    List.Add 2, "b"
    List.Add 3, "c"

    Dim Actual As IList
    Set Actual = List.GetKeyList
    
    Dim Expected() As Variant
    Expected = Array(1, 2, 3)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub KeyList_WithIndex_ReturnsValue()
    Dim List As New SortedList
    List.Add 1, "a"
    List.Add 2, "b"
    List.Add 3, "c"

    Dim Actual As IList
    Set Actual = List.GetKeyList

    Assert.That Actual(1), Iz.EqualTo(2)
End Sub

Public Sub GetValueList_WhenCalled_ReturnsListOfValues()
    Dim List As New SortedList
    List.Add 1, "a"
    List.Add 2, "b"
    List.Add 3, "c"

    Dim Actual As IList
    Set Actual = List.GetValueList
    
    Dim Expected() As Variant
    Expected = Array("a", "b", "c")
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub NewSortedList_WithHashtable_CreatesFilledList()
    Dim Table As New Hashtable
    Table.Add 1, "a"
    Table.Add 3, "c"
    Table.Add 2, "b"
    
    Dim Actual As SortedList
    Set Actual = Cor.NewSortedList(Table)
    
    Dim Expected() As Variant
    Expected = Array(Cor.NewDictionaryEntry(1, "a"), Cor.NewDictionaryEntry(2, "b"), Cor.NewDictionaryEntry(3, "c"))
    Assert.That Actual, Iz.EqualTo(Expected).Using(New CorObjectComparer)
End Sub
