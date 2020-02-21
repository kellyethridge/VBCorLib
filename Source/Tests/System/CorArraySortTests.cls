VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorArraySortTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CorArraySortTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Sort_WithPrimitives_SortsAscending"
            Test.Use NewBytes(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewIntegers(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewLongs(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewSingles(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewDoubles(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewCurrencies(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewStrings("c", "d", "a", "B", "A")
            Test.Use NewDates(#5/1/2001#, #4/1/2001#, #6/1/2001#, #3/1/2001#)
            Test.Use NewVariants(7, 8, 9, 6, 5, 4, 1, 3, 2)
            Test.Use NewInt64s(7, 8, 9, 6, 5, 4, 1, 3, 2)
            
        Case "Sort_WithCallback_SortsItems"
            Test.Use NewBytes(7, 8, 9, 6, 5, 4, 1, 3, 2), AddressOf CompareBytes
            Test.Use NewIntegers(7, 8, 9, 6, 5, 4, 1, 3, 2), AddressOf CompareIntegers
            Test.Use NewLongs(7, 8, 9, 6, 5, 4, 1, 3, 2), AddressOf CompareLongs
            Test.Use NewSingles(7, 8, 9, 6, 5, 4, 1, 3, 2), AddressOf CompareSingles
            Test.Use NewDoubles(7, 8, 9, 6, 5, 4, 1, 3, 2), AddressOf CompareDoubles
            Test.Use NewCurrencies(7, 8, 9, 6, 5, 4, 1, 3, 2), AddressOf CompareCurrencies
            Test.Use NewStrings("c", "d", "a", "B", "A"), AddressOf CompareStrings
            Test.Use NewDates(#5/1/2001#, #4/1/2001#, #6/1/2001#, #3/1/2001#), AddressOf CompareDates
            Test.Use NewVariants(7, 8, 9, 6, 5, 4, 1, 3, 2), AddressOf CompareVariants
            
    End Select
End Sub

Public Sub Sort_WithPrimitives_SortsAscending(ByRef Arr As Variant)
    CorArray.Sort Arr
    
    Assert.That Arr, Iz.Ordered.Using(New CorObjectComparer)
End Sub

Public Sub Sort_WithNullArray_ThrowsArgumentNullException()
    Dim Arr() As Long
    
    On Error Resume Next
    CorArray.Sort Arr
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub Sort_WithMultiDimensionArray_ThrowsRankException()
    Dim Arr() As Long
    ReDim Arr(1, 1)
    
    On Error Resume Next
    CorArray.Sort Arr
    
    AssertRankException Err
End Sub

Public Sub Sort_WithInvalidComparer_ThrowsArgumentException()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.Sort Arr, "Comparer"
    
    AssertArgumentException Err, "Comparer"
End Sub

Public Sub Sort_WithComparer_SortsValues()
    Dim Arr() As Long
    Arr = NewLongs(7, 8, 9, 6, 5, 4, 1, 3, 2)
    
    CorArray.Sort Arr, New DescendingComparer
    
    Assert.That Arr, Iz.Ordered.Descending
End Sub

Public Sub Sort_WithComparableObjects_SortsObjects()
    Dim a(4) As Int32
    Set a(0) = NewInt32(3)
    Set a(1) = NewInt32(2)
    Set a(2) = NewInt32(5)
    Set a(3) = NewInt32(4)
    Set a(4) = NewInt32(0)
    
    CorArray.Sort a
    
    Assert.That a, Iz.Ordered.Using(New ComparableComparer)
End Sub

Public Sub Sort_WithCallback_SortsItems(ByRef Items As Variant, ByVal Callback As Long)
    CorArray.Sort Items, Callback
    
    Assert.That Items, Iz.Ordered
End Sub

Public Sub Sort_WithUDTsUsingCallback_SortsUDTs()
    Dim a(4) As VBGUID
    a(0).Data1 = 3
    a(1).Data1 = 2
    a(2).Data1 = 4
    a(3).Data1 = 1
    a(4).Data1 = 6
    
    CorArray.Sort a, AddressOf CompareVBGuids
    
    Assert.That a, Iz.Ordered.Using(New VBGuidComparer)
End Sub

Public Sub Sort_LargeUdtsUsingCallback_SortsUdts()
    Dim Udts(4) As LargeUdt
    Udts(0).Value = 3
    Udts(1).Value = 2
    Udts(2).Value = 4
    Udts(3).Value = 1
    Udts(4).Value = 6
    
    CorArray.Sort Udts, AddressOf CompareLargeUdts
    
    Assert.That Udts, Iz.Ordered.Using(New LargeUdtComparer)
End Sub

Public Sub Sort_WithNonZeroLBound_SortsValues()
    Dim Arr(5 To 9) As Long
    Arr(5) = 5
    Arr(6) = 2
    Arr(7) = 3
    Arr(8) = 6
    Arr(9) = 1
    
    CorArray.Sort Arr
    
    Assert.That Arr, Iz.Ordered
End Sub

Public Sub SortEx_WithNullArray_Throws()
    Dim a() As Long
    
    On Error Resume Next
    CorArray.SortEx a, 2, 4
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub SortEx_WithMultiDimensionArray_Throws()
    Dim Arr(1, 1) As Long
    
    On Error Resume Next
    CorArray.SortEx Arr, 0, 1
    
    AssertRankException Err
End Sub

Public Sub SortEx_WithIndexLessThanLBound_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.SortEx Arr, -1, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub SortEx_WithLengthLessThanZero_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.SortEx Arr, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub SortEx_WithRangeExtendingPassedEndOfArray_Throws()
    Dim Arr(1 To 10) As Long
    
    On Error Resume Next
    CorArray.SortEx Arr, 1, 11
    
    AssertArgumentException Err
End Sub

Public Sub SortEx_WhenCalled_SortsSubArray()
    Dim a() As Long
    a = NewLongs(6, 5, 4, 1, 3, 2, 1, 0)
    
    CorArray.SortEx a, 2, 4
    
    Dim Expected() As Long
    Expected = NewLongs(6, 5, 1, 2, 3, 4, 1, 0)
    Assert.That a, Iz.EqualTo(Expected)
End Sub

Public Sub SortEx_WithComparer_SortsSubArray()
    Dim a() As Long
    a = NewLongs(5, 6, 4, 1, 3, 2, 1, 0)
    
    CorArray.SortEx a, 2, 4, New DescendingComparer
    
    Dim Expected() As Long
    Expected = NewLongs(5, 6, 4, 3, 2, 1, 1, 0)
    Assert.That a, Iz.EqualTo(Expected)
End Sub

Public Sub SortKey_WithNullKeysArray_Throws()
    Dim Keys() As Long
    Dim Items(0) As Long
    
    On Error Resume Next
    CorArray.SortKey Keys, Items
    
    AssertArgumentNullException Err, "Keys"
End Sub

Public Sub SortKey_WithNullItemsArray_Throws()
    Dim Keys(0) As Long
    Dim Items() As Long
    
    On Error Resume Next
    CorArray.SortKey Keys, Items
    
    AssertArgumentNullException Err, "Items"
End Sub

Public Sub SortKey_WithMultiDimensionKeys_Throws()
    Dim Keys(0, 0) As Long
    Dim Items(0) As Long
    
    On Error Resume Next
    CorArray.SortKey Keys, Items
    
    AssertRankException Err
End Sub

Public Sub SortKey_WithMultiDimensionItems_Throws()
    Dim Keys(0) As Long
    Dim Items(0, 0) As Long
    
    On Error Resume Next
    CorArray.SortKey Keys, Items
    
    AssertRankException Err
End Sub

Public Sub SortKey_WithInvalidItems_Throws()
    Dim Keys(0) As Long
    Dim InvalidItems As Long
    
    On Error Resume Next
    CorArray.SortKey Keys, InvalidItems
    
    AssertArgumentException Err, "Items"
End Sub

Public Sub SortKeyEx_WithNullKeys_Throws()
    Dim Keys() As Long
    Dim Items(0) As Long
    
    On Error Resume Next
    CorArray.SortKeyEx Keys, Items, 0, 1
    
    AssertArgumentNullException Err, "Keys"
End Sub

Public Sub SortKey_WhenCalled_SortsKeys()
    Dim Keys(4) As Long
    Dim Items(4) As Long
    Keys(0) = 3: Items(0) = 1
    Keys(1) = 5: Items(1) = 2
    Keys(2) = 2: Items(2) = 3
    Keys(3) = 6: Items(3) = 4
    Keys(4) = 1: Items(4) = 5
    
    CorArray.SortKey Keys, Items
    
    Assert.That Keys, Iz.Ordered
End Sub

Public Sub SortKey_WhenCalled_SortsItemsByKeys()
    Dim Keys(4) As Long
    Dim Items(4) As Long
    Keys(0) = 3: Items(0) = 1
    Keys(1) = 5: Items(1) = 2
    Keys(2) = 2: Items(2) = 3
    Keys(3) = 6: Items(3) = 4
    Keys(4) = 1: Items(4) = 5
    
    CorArray.SortKey Keys, Items
    
    Dim Expected() As Long
    Expected = NewLongs(5, 3, 1, 2, 4)
    Assert.That Items, Iz.EqualTo(Expected)
End Sub

Public Sub SortKey_WithComparer_SortsKeys()
    Dim Keys(4) As Long
    Dim Items(4) As Long
    Keys(0) = 3: Items(0) = 1
    Keys(1) = 5: Items(1) = 2
    Keys(2) = 2: Items(2) = 3
    Keys(3) = 6: Items(3) = 4
    Keys(4) = 1: Items(4) = 5
    
    CorArray.SortKey Keys, Items, New DescendingComparer
    
    Assert.That Keys, Iz.Ordered.Descending
End Sub

Public Sub SortKey_WithComparer_SortsItemsByKeys()
    Dim Keys(4) As Long
    Dim Items(4) As Long
    Keys(0) = 3: Items(0) = 1
    Keys(1) = 5: Items(1) = 2
    Keys(2) = 2: Items(2) = 3
    Keys(3) = 6: Items(3) = 4
    Keys(4) = 1: Items(4) = 5
    
    CorArray.SortKey Keys, Items, New DescendingComparer
    
    Dim Expected() As Long
    Expected = NewLongs(4, 2, 1, 3, 5)
    Assert.That Items, Iz.EqualTo(Expected)
End Sub



