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
Option Explicit
Implements ITestCaseSource

Private original As ArrayList
Private list As ArrayList

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

Public Sub Setup()
    Set original = New ArrayList
    Dim i As Long
    For i = 1 To 20
        original.Add i
    Next i
    Set list = original.GetRange(4, 0)
End Sub

Public Sub GetRange_WithCountOfZero_ReturnsEmptyList()
    Dim list As New ArrayList
    Dim Ranged As ArrayList
    
    Set Ranged = list.GetRange(0, 0)
    
    Assert.That Ranged, Iz.Empty
End Sub

Public Sub Add_WhenCalled_AddsToUnderlyingList()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Ranged.Add 1
    
    Assert.That Ranged, Iz.EqualTo(NewLongs(3, 4, 1, 2, 3, 1))
    Assert.That list, Iz.EqualTo(NewLongs(1, 2, 3, 4, 1, 2, 3, 1, 4))
End Sub

Public Sub Item_WhenCalled_ReturnsItemAtIndex()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Assert.That Ranged(0), Iz.EqualTo(3)
    Assert.That Ranged(2), Iz.EqualTo(1)
    Assert.That Ranged(4), Iz.EqualTo(3)
End Sub

Public Sub AddRange_WithCollection_AppendsItems()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    Dim Source As New Collection
    Source.Add 1
    Source.Add 2
    Source.Add 3
    
    Ranged.AddRange Source
    
    Assert.That Ranged, Iz.EqualTo(NewLongs(3, 4, 1, 2, 3, 1, 2, 3))
    Assert.That list, Iz.EqualTo(NewLongs(1, 2, 3, 4, 1, 2, 3, 1, 2, 3, 4))
End Sub

Public Sub AddRange_WithArray_AppendsItems()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    Dim Source() As Long
    Source = NewLongs(1, 2, 3)
    
    Ranged.AddRange Source
    
    Assert.That Ranged, Iz.EqualTo(NewLongs(3, 4, 1, 2, 3, 1, 2, 3))
    Assert.That list, Iz.EqualTo(NewLongs(1, 2, 3, 4, 1, 2, 3, 1, 2, 3, 4))
End Sub

Public Sub Item_WhenSet_SetsValueInUnderlyingList()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Ranged(2) = "Hello"
    
    Assert.That Ranged, Iz.EqualTo(Array(3, 4, "Hello", 2, 3))
    Assert.That list, Iz.EqualTo(Array(1, 2, 3, 4, "Hello", 2, 3, 4))
End Sub

Public Sub Clear_WhenCalled_RemovesRangedItemsFromList()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Ranged.Clear
    
    Assert.That Ranged, Iz.Empty
    Assert.That list, Iz.EqualTo(NewLongs(1, 2, 4))
End Sub

Public Sub CopyTo_WhenCalled_CopiesRangedItems()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    Dim Target(0 To 4) As String
    
    Ranged.CopyTo Target
    
    Assert.That Target, Iz.EqualTo(Array("3", "4", "1", "2", "3"))
End Sub

Public Sub CopyTo_WithArrayIndex_CopiesRangedItems()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    Dim Target(0 To 5) As String
    
    Ranged.CopyTo Target, 1
    
    Assert.That Target, Iz.EqualTo(Array("", "3", "4", "1", "2", "3"))
End Sub

Public Sub GetEnumerator_WhenCalled_EnumeratesItems()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Dim Actual As Object
    Set Actual = Ranged.GetEnumerator
    
    Assert.That Actual, Iz.EqualTo(Array(3, 4, 1, 2, 3))
End Sub

Public Sub GetEnumerator_WithSubset_EnumeratesSubset()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)

    Dim Actual As Object
    Set Actual = Ranged.GetEnumerator(1, 3)
    
    Assert.That Actual, Iz.EqualTo(Array(4, 1, 2))
End Sub

Public Sub IndexOf_WhenCalled_ReturnsExpected()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Dim Actual As Long
    Actual = Ranged.IndexOf(4&)
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub IndexOf_WithNegativeStartIndex_ThrowsArgumentOutOfRangeException()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    On Error Resume Next
    Ranged.IndexOf 4&, -1, 1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub IndexOf_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    On Error Resume Next
    Ranged.IndexOf 4&, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub IndexOf_WithInvalidRange_ThrowsArgumentException()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    On Error Resume Next
    Ranged.IndexOf 4&, 0, 6
    
    AssertArgumentException Err, ""
End Sub

Public Sub Insert_WithValues_InsertsIntoUnderlyingList(ByVal Index As Long, ByVal Value As Variant, ByVal Expected As Variant)
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)

    Ranged.Insert Index, Value
    
    Assert.That list, Iz.EqualTo(Expected)
End Sub

Public Sub InsertRange_WithArray_InsertsItems()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Ranged.Insertrange 2, NewStrings("Item1", "Item2", "Item3")
    
    Assert.That list, Iz.EqualTo(Array(1, 2, 3, 4, "Item1", "Item2", "Item3", 1, 2, 3, 4))
End Sub

Public Sub InsertRange_WithCollection_InsertsItems()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Ranged.Insertrange 2, NewCollection("Item1", "Item2", "Item3")
    
    Assert.That list, Iz.EqualTo(Array(1, 2, 3, 4, "Item1", "Item2", "Item3", 1, 2, 3, 4))
End Sub

Public Sub LastIndexOf_WithValue_ReturnsExpected(ByVal Value As Long, ByVal Expected As Long)
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Dim Actual As Long
    Actual = Ranged.LastIndexOf(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub RemoveAt_WithIndexZero_RemovesFirstElement()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Ranged.RemoveAt 0
    
    Assert.That list, Iz.EqualTo(Array(1, 2, 4, 1, 2, 3, 4))
End Sub

Public Sub Remove_WithExistingValue_RemovesFirstInstance()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Ranged.Remove 3&
    
    Assert.That list, Iz.EqualTo(Array(1, 2, 4, 1, 2, 3, 4))
End Sub

Public Sub RemoveRange_WhenCalled_RemovesElements()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Ranged.RemoveRange 1, 2
    
    Assert.That list, Iz.EqualTo(Array(1, 2, 3, 2, 3, 4))
End Sub

Public Sub Reverse_WhenCalled_ReversedElements()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    Ranged.Reverse
    
    Assert.That list, Iz.EqualTo(Array(1, 2, 3, 2, 1, 4, 3, 4))
End Sub

Public Sub ToArray_WhenCalled_CopiesRangeToArray()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)
    
    
    Dim Actual() As Long
    Actual = Ranged.ToArray(vbLong)
    
    Assert.That Actual, Iz.EqualTo(Array(3, 4, 1, 2, 3))
End Sub

Public Sub SetRanged_WhenCalled_OverritesValuesInRange()
    Dim list As ArrayList
    Dim Ranged As ArrayList
    Set Ranged = MakeFilledRanged(list)

    Ranged.SetRange 0, Array("a", "b")
    
    Assert.That list, Iz.EqualTo(Array(1, 2, "a", "b", 1, 2, 3, 4))
End Sub


Private Function MakeFilledRanged(ByRef list As ArrayList) As ArrayList
    Set list = New ArrayList
    list.AddRange (NewLongs(1, 2, 3, 4, 1, 2, 3, 4))
    Set MakeFilledRanged = list.GetRange(2, 5)
End Function

