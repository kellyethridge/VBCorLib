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

Private original As ArrayList
Private List As ArrayList

Public Sub Setup()
    Set original = New ArrayList
    Dim i As Long
    For i = 1 To 20
        original.Add i
    Next i
    Set List = original.GetRange(4, 0)
End Sub

Public Sub GetRange_WithCountOfZero_ReturnsEmptyList()
    Assert.AreEqual 0, List.Count
    AssertOriginalList
End Sub

Public Sub TestAfterAddingOne()
    List.Add 1
    Assert.AreEqual 1, List.Count
    AssertOriginalList
End Sub

Public Sub TestAfterAddingTwo()
    List.Add 1
    List.Add 2
    Assert.AreEqual 2, List.Count
    AssertOriginalList
End Sub

Public Sub TestRetrieveAfterAdding()
    List.Add 1
    Assert.AreEqual 1, List(0)
    AssertOriginalList
End Sub

Public Sub TestAddRangeFromCollection()
    Dim col As New Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    List.AddRange col
    Assert.AreEqual 3, List.Count
    Assert.AreEqual 1, List(0)
    Assert.AreEqual 2, List(1)
    Assert.AreEqual 3, List(2)
    AssertOriginalList
End Sub

Public Sub TestAddRangeFromArray()
    Dim a() As Long
    a = NewLongs(1, 2, 3)
    
    List.AddRange a
    Assert.AreEqual 3, List.Count
    Assert.AreEqual 1, List(0)
    Assert.AreEqual 2, List(1)
    Assert.AreEqual 3, List(2)
    AssertOriginalList
End Sub

Public Sub TestSetItem()
    List.AddRange NewLongs(1, 2, 3, 4, 5)
    List(2) = "Hello"
    Assert.AreEqual 5, List.Count
    Assert.AreEqual 1, List(0)
    Assert.AreEqual 2, List(1)
    Assert.AreEqual "Hello", List(2)
    Assert.AreEqual 4, List(3)
    Assert.AreEqual 5, List(4)
    AssertOriginalList
End Sub

Public Sub TestClear()
    List.Add 1
    List.Clear
    Assert.AreEqual 0, List.Count
    AssertOriginalList
End Sub

Public Sub TestCopyTo()
    Dim a(2) As String
    List.Add "Item1"
    List.Add "Item2"
    List.Add "Item3"
    List.CopyTo a
    Assert.AreEqual "Item1", a(0)
    Assert.AreEqual "Item2", a(1)
    Assert.AreEqual "Item3", a(2)
End Sub

Public Sub TestCopyToIndexOffset()
    Dim a(5) As String
    List.Add "Item1"
    List.Add "Item2"
    List.Add "Item3"
    List.CopyTo a, 2
    Assert.AreEqual "Item1", a(2)
    Assert.AreEqual "Item2", a(3)
    Assert.AreEqual "Item3", a(4)
End Sub

Public Sub TestEnumerator()
    List.AddRange NewStrings("Item1", "Item2", "Item3")
    
    Dim v As Variant
    For Each v In List
        Assert.AreEqual "Item", Left$(v, 4)
    Next v
End Sub

Public Sub TestEnumeratorIndex()
    Dim i As Long
    For i = 1 To 10
        List.Add i
    Next i
    
    Dim v As Variant
    i = 3
    For Each v In List.GetEnumerator(2, 5)
        Assert.AreEqual i, v
        i = i + 1
    Next v
End Sub

Public Sub TestIndexOf()
    Dim i As Long
    For i = 1 To 10
        List.Add "Item " & i
    Next i
    Assert.AreEqual 8, List.IndexOf("Item 9")
End Sub

Public Sub TestInsertAtBeginning()
    Dim i As Long
    For i = 1 To 10
        List.Add "Item " & i
    Next i
    List.Insert 0, 1
    Assert.AreEqual 11, List.Count
    Assert.AreEqual 1, List(0)
    For i = 1 To 10
        Assert.AreEqual "Item " & i, List(i)
    Next i
    AssertOriginalList
End Sub

Public Sub TestInsertAtEnd()
    Dim i As Long
    For i = 1 To 10
        List.Add "Item " & i
    Next i
    List.Insert 10, 1
    Assert.AreEqual 1, List(10)
    For i = 1 To 10
        Assert.AreEqual "Item " & i, List(i - 1)
    Next i
    AssertOriginalList
End Sub

Public Sub TestInsertRangeArray()
    List.AddRange NewLongs(1, 2, 3, 4)
    List.Insertrange 2, NewStrings("Item1", "Item2", "Item3")
    AssertInsert
    AssertOriginalList
End Sub

Public Sub TestInsertRangecollection()
    List.AddRange NewLongs(1, 2, 3, 4)
    Dim col As New Collection
    col.Add "Item1"
    col.Add "Item2"
    col.Add "Item3"
    List.Insertrange 2, col
    AssertInsert
    AssertOriginalList
End Sub

Public Sub TestLastIndexOf()
    Dim i As Long
    For i = 1 To 10
        List.Add "Item " & i
    Next i
    
    Assert.AreEqual 3, List.LastIndexOf("Item 4"), "LastIndexOf should return 3."
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


