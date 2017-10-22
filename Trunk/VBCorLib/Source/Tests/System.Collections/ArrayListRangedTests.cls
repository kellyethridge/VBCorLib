VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayListRangedTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ArrayListRangedTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Insert_WithValues_InsertsIntoUnderlyingList"
            Test.Use 0, "a", Array(1, 2, "a", 3, 4, 1, 2, 3, 4)
            Test.Use 5, "a", Array(1, 2, 3, 4, 1, 2, 3, "a", 4)
        
        Case "LastIndexOf_WithValue_ReturnsExpected"
            Test.Use 4, 1
            Test.Use 3, 4
    End Select
End Sub

Public Sub GetRange_WithCountOfZero_ReturnsEmptyList()
    Dim List As New ArrayList
    Dim Actual As ArrayList
    List.AddRange NewLongs(1, 2, 3, 4)
    
    Set Actual = List.GetRange(0, 0)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub GetRange_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim List As New ArrayList
    On Error Resume Next
    
    List.GetRange -1, 0
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetRange_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim List As New ArrayList
    On Error Resume Next
    
    List.GetRange 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub GetRange_WithIndexPlusCountBeyondListSize_ThrowsArgumentException()
    Dim List As ArrayList
    Set List = ArrayList.Repeat(1, 10)
    On Error Resume Next
    
    List.GetRange 2, 9
    
    AssertArgumentException Err
End Sub

Public Sub Add_WhenCalled_AddsToUnderlyingList()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Ranged.Add 1
    
    Assert.That Ranged, Iz.EqualTo(NewLongs(3, 4, 1, 2, 3, 1))
    Assert.That List, Iz.EqualTo(NewLongs(1, 2, 3, 4, 1, 2, 3, 1, 4))
End Sub

Public Sub Item_WhenCalled_ReturnsItemAtIndex()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Assert.That Ranged(0), Iz.EqualTo(3)
    Assert.That Ranged(2), Iz.EqualTo(1)
    Assert.That Ranged(4), Iz.EqualTo(3)
End Sub

Public Sub AddRange_WithCollection_AppendsItems()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    Dim Source As New Collection
    Source.Add 1
    Source.Add 2
    Source.Add 3
    
    Ranged.AddRange Source
    
    Assert.That Ranged, Iz.EqualTo(NewLongs(3, 4, 1, 2, 3, 1, 2, 3))
    Assert.That List, Iz.EqualTo(NewLongs(1, 2, 3, 4, 1, 2, 3, 1, 2, 3, 4))
End Sub

Public Sub AddRange_WithArray_AppendsItems()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    Dim Source() As Long
    Source = NewLongs(1, 2, 3)
    
    Ranged.AddRange Source
    
    Assert.That Ranged, Iz.EqualTo(NewLongs(3, 4, 1, 2, 3, 1, 2, 3))
    Assert.That List, Iz.EqualTo(NewLongs(1, 2, 3, 4, 1, 2, 3, 1, 2, 3, 4))
End Sub

Public Sub Item_WhenSet_SetsValueInUnderlyingList()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Ranged(2) = "Hello"
    
    Assert.That Ranged, Iz.EqualTo(Array(3, 4, "Hello", 2, 3))
    Assert.That List, Iz.EqualTo(Array(1, 2, 3, 4, "Hello", 2, 3, 4))
End Sub

Public Sub Clear_WhenCalled_RemovesRangedItemsFromList()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Ranged.Clear
    
    Assert.That Ranged, Iz.Empty
    Assert.That List, Iz.EqualTo(NewLongs(1, 2, 4))
End Sub

Public Sub CopyTo_WhenCalled_CopiesRangedItems()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    Dim Target(0 To 4) As String
    
    Ranged.CopyTo Target
    
    Assert.That Target, Iz.EqualTo(Array("3", "4", "1", "2", "3"))
End Sub

Public Sub CopyTo_WithArrayIndex_CopiesRangedItems()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    Dim Target(0 To 5) As String
    
    Ranged.CopyTo Target, 1
    
    Assert.That Target, Iz.EqualTo(Array("", "3", "4", "1", "2", "3"))
End Sub

Public Sub GetEnumerator_WhenCalled_EnumeratesItems()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Dim Actual As Object
    Set Actual = Ranged.GetEnumerator
    
    Assert.That Actual, Iz.EqualTo(Array(3, 4, 1, 2, 3))
End Sub

Public Sub GetEnumerator_WithSubset_EnumeratesSubset()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)

    Dim Actual As Object
    Set Actual = Ranged.GetEnumerator(1, 3)
    
    Assert.That Actual, Iz.EqualTo(Array(4, 1, 2))
End Sub

Public Sub IndexOf_WhenCalled_ReturnsExpected()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Dim Actual As Long
    Actual = Ranged.IndexOf(4&)
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub IndexOf_WithNegativeStartIndex_ThrowsArgumentOutOfRangeException()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    On Error Resume Next
    Ranged.IndexOf 4&, -1, 1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub IndexOf_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    On Error Resume Next
    Ranged.IndexOf 4&, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub IndexOf_WithInvalidRange_ThrowsArgumentException()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    On Error Resume Next
    Ranged.IndexOf 4&, 0, 6
    
    AssertArgumentException Err, ""
End Sub

Public Sub Insert_WithValues_InsertsIntoUnderlyingList(ByVal Index As Long, ByVal Value As Variant, ByVal Expected As Variant)
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)

    Ranged.Insert Index, Value
    
    Assert.That List, Iz.EqualTo(Expected)
End Sub

Public Sub InsertRange_WithArray_InsertsItems()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Ranged.Insertrange 2, NewStrings("Item1", "Item2", "Item3")
    
    Assert.That List, Iz.EqualTo(Array(1, 2, 3, 4, "Item1", "Item2", "Item3", 1, 2, 3, 4))
End Sub

Public Sub InsertRange_WithCollection_InsertsItems()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Ranged.Insertrange 2, NewCollection("Item1", "Item2", "Item3")
    
    Assert.That List, Iz.EqualTo(Array(1, 2, 3, 4, "Item1", "Item2", "Item3", 1, 2, 3, 4))
End Sub

Public Sub LastIndexOf_WithValue_ReturnsExpected(ByVal Value As Long, ByVal Expected As Long)
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Dim Actual As Long
    Actual = Ranged.LastIndexOf(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub LastIndexOf_WithMissingCount_ThrowsArgumentException()
    Dim List As ArrayList
    Set List = MakeRangedList()
    
    On Error Resume Next
    List.LastIndexOf 1, 1
    
    AssertArgumentException Err, "Count"
End Sub

Public Sub LastIndexOf_WithMissingStartIndex_ThrowsArgumentException()
    Dim List As ArrayList
    Set List = MakeRangedList()
    
    On Error Resume Next
    List.LastIndexOf 1, , 1
    
    AssertArgumentException Err, "StartIndex"
End Sub

Public Sub LastIndexOf_WithStartIndexPastEndOfList_ThrowsArgumentOutOfRangeException()
    Dim List As ArrayList
    Set List = MakeRangedList
    
    On Error Resume Next
    List.LastIndexOf 1, List.Count, 1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub LastIndexOf_WithNegativeCount_ThrowArgumentOutOfRangeException()
    Dim List As ArrayList
    Set List = MakeRangedList
    
    On Error Resume Next
    List.LastIndexOf 1, List.Count - 1, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub LastIndexOf_WithInvalidPositionWithinList_ThrowsArgumentException()
    Dim List As ArrayList
    Set List = MakeRangedList
    
    On Error Resume Next
    List.LastIndexOf 1, List.Count - 1, List.Count + 1
    
    AssertArgumentException Err, "Count"
End Sub

Public Sub RemoveAt_WithIndexZero_RemovesFirstElement()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Ranged.RemoveAt 0
    
    Assert.That List, Iz.EqualTo(Array(1, 2, 4, 1, 2, 3, 4))
End Sub

Public Sub Remove_WithExistingValue_RemovesFirstInstance()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Ranged.Remove 3&
    
    Assert.That List, Iz.EqualTo(Array(1, 2, 4, 1, 2, 3, 4))
End Sub

Public Sub RemoveRange_WhenCalled_RemovesElements()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Ranged.RemoveRange 1, 2
    
    Assert.That List, Iz.EqualTo(Array(1, 2, 3, 2, 3, 4))
End Sub

Public Sub Reverse_WhenCalled_ReversedElements()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    Ranged.Reverse
    
    Assert.That List, Iz.EqualTo(Array(1, 2, 3, 2, 1, 4, 3, 4))
End Sub

Public Sub Reverse_WithNegativeIndex_ThrowsArgumentOutOfRangeException()
    Dim List As ArrayList
    Set List = MakeRangedList
    
    On Error Resume Next
    List.Reverse -1, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Reverse_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    Dim List As ArrayList
    Set List = MakeRangedList
    
    On Error Resume Next
    List.Reverse 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Reverse_WithIndexAndCountExceedLength_ThrowsArgumentException()
    Dim List As ArrayList
    Set List = MakeRangedList
    
    On Error Resume Next
    List.Reverse 5, 6
    
    AssertArgumentException Err
End Sub

Public Sub Reverse_WithOnlyIndex_ThrowsArgumentException()
    Dim List As ArrayList
    Set List = MakeRangedList
    
    On Error Resume Next
    List.Reverse 1
    
    AssertArgumentException Err, "Count"
End Sub

Public Sub Reverse_WithOnlyCount_ThrowsArgumentException()
    Dim List As ArrayList
    Set List = MakeRangedList
    
    On Error Resume Next
    List.Reverse , 1
    
    AssertArgumentException Err, "Index"
End Sub

Public Sub ToArray_WhenCalled_CopiesRangeToArray()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)
    
    
    Dim Actual() As Long
    Actual = Ranged.ToArray(vbLong)
    
    Assert.That Actual, Iz.EqualTo(Array(3, 4, 1, 2, 3))
End Sub

Public Sub SetRanged_WhenCalled_OverritesValuesInRange()
    Dim List As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(List)

    Ranged.SetRange 0, Array("a", "b")
    
    Assert.That List, Iz.EqualTo(Array(1, 2, "a", "b", 1, 2, 3, 4))
End Sub

Private Function MakeRangedList() As ArrayList
    Set MakeRangedList = MakeFilledRanged(New ArrayList)
End Function

Private Function MakeFilledRanged(ByRef List As ArrayList) As ArrayList
    Set List = New ArrayList
    List.AddRange NewLongs(1, 2, 3, 4, 1, 2, 3, 4)
    Set MakeFilledRanged = List.GetRange(2, 5)
End Function

