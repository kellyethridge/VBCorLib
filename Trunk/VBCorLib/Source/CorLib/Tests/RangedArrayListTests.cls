VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangedArrayListTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private original As ArrayList
Private List As ArrayList

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Insert_WithValues_InsertsIntoUnderlyingList"
            Test.Use 0, "a", Array(1, 2, "a", 3, 4, 1, 2, 3, 4)
            Test.Use 5, "a", Array(1, 2, 3, 4, 1, 2, 3, "a", 4)
            
    End Select
End Sub

Public Sub Setup()
    Set original = New ArrayList
    Dim i As Long
    For i = 1 To 20
        original.Add i
    Next i
    Set List = original.GetRange(4, 0)
End Sub

Public Sub GetRange_WithCountOfZero_ReturnsEmptyList()
    Dim List As New ArrayList
    Dim Ranged As ArrayList
    
    Set Ranged = List.GetRange(0, 0)
    
    Assert.That Ranged, Iz.Empty
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

End Sub

Public Sub TestRemoveAtBeginning()
    Dim i As Long
    For i = 1 To 10
        List.Add "Item " & i
    Next i
        
    List.RemoveAt 0
    Assert.AreEqual 9, List.Count, "count should return 9."
    For i = 0 To 8
        Assert.AreEqual "Item " & i + 2, List(i), "list should return 'Item " & i + 2 & "'."
    Next i
    AssertOriginalList
End Sub

Public Sub TestRemoveAtEnd()
    Dim i As Long
    For i = 1 To 10
        List.Add "Item " & i
    Next i
    
    List.RemoveAt 9
    Assert.AreEqual 9, List.Count, "Count should return 9."
    For i = 0 To 8
        Assert.AreEqual "Item " & i + 1, List(i), "list should return 'Item " & i + 1 & "'."
    Next i
    AssertOriginalList
End Sub

Public Sub TestRemove()
    Dim i As Long
    For i = 1 To 10
        List.Add "Item " & i
    Next i
    
    List.Remove "Item 3"
    Assert.AreEqual 9, List.Count, "Count should return 9."
    Assert.AreEqual "Item 1", List(0), "list should return 'Item 1'."
    Assert.AreEqual "Item 2", List(1), "list should return 'Item 2'."
    For i = 2 To 8
        Assert.AreEqual "Item " & i + 2, List(i), "list should return 'Item " & i + 2 & "'."
    Next i
    AssertOriginalList
End Sub

Public Sub TestRemoveRange()
    Dim i As Long
    For i = 1 To 10
        List.Add "Item " & i
    Next i

    List.RemoveRange 2, 4
    Assert.AreEqual 6, List.Count, "Count should return 6."
    Assert.AreEqual "Item 1", List(0), "list should return 'Item 1'."
    Assert.AreEqual "Item 2", List(1), "list should return 'Item 2'."
    For i = 2 To 5
        Assert.AreEqual "Item " & i + 5, List(i), "list should return 'Item " & i + 5 & "'."
    Next i
    AssertOriginalList
End Sub

Public Sub TestReverse()
    Dim i As Long
    For i = 1 To 10
        List.Add i
    Next i
    List.Reverse
    For i = 0 To 9
        Assert.AreEqual 10 - i, List(i), "list should return " & 10 - i
    Next i
    AssertOriginalList
End Sub

Public Sub TestToArray()
    Dim i As Long
    For i = 1 To 10
        List.Add i
    Next i
    Dim a() As Long
    a = List.ToArray(vbLong)
    For i = 0 To 9
        Assert.AreEqual i + 1, a(i)
    Next i
End Sub

Public Sub TestSetRange()
    Dim i As Long
    For i = 1 To 10
        List.Add i
    Next i
    List.SetRange 2, NewStrings("Item 3", "Item 4", "Item 5")
    Assert.AreEqual 10, List.Count, "Count should return 10."
    For i = 0 To 9
        If i >= 2 And i <= 4 Then
            Assert.AreEqual "Item " & i + 1, List(i), "list should return 'Item " & i - 1 & "'."
        Else
            Assert.AreEqual i + 1, List(i)
        End If
    Next i
    AssertOriginalList
End Sub


Private Function MakeFilledRanged(ByRef List As ArrayList) As ArrayList
    Set List = New ArrayList
    List.AddRange (NewLongs(1, 2, 3, 4, 1, 2, 3, 4))
    Set MakeFilledRanged = List.GetRange(2, 5)
End Function

Private Sub AssertInsert()
    Assert.AreEqual 7, List.Count, "Count should return 7."
    Assert.AreEqual 1, List(0), "list should return '1'."
    Assert.AreEqual 2, List(1), "list should return '2'."
    Assert.AreEqual "Item1", List(2), "list should return 'Item1'."
    Assert.AreEqual "Item2", List(3), "list should return 'Item2'."
    Assert.AreEqual "Item3", List(4), "list should return 'Item3'."
    Assert.AreEqual 3, List(5), "list should return '3'."
    Assert.AreEqual 4, List(6), "list should return '4'."
End Sub


Private Sub AssertOriginalList()
    Assert.AreEqual 20 + List.Count, original.Count, "Original count should be " & 20 + List.Count
    
    Dim i As Long
    For i = 0 To 3
        Assert.AreEqual i + 1, original(i), "original should return '" & i & "'."
    Next i
    For i = 4 + List.Count To 19 + List.Count
        Assert.AreEqual i - List.Count + 1, original(i), "original should return '" & i - List.Count + 1 & "'."
    Next i
End Sub


