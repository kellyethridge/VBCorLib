VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayListRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private original As ArrayList
Private list As ArrayList

Public Sub Setup()
    Set original = New ArrayList
    Dim i As Long
    For i = 1 To 20
        original.Add i
    Next i
    Set list = original.GetRange(4, 0)
End Sub

Public Sub TestEmptyList()
    Assert.AreEqual 0, list.Count
    AssertOriginalList
End Sub

Public Sub TestAfterAddingOne()
    list.Add 1
    Assert.AreEqual 1, list.Count
    AssertOriginalList
End Sub

Public Sub TestAfterAddingTwo()
    list.Add 1
    list.Add 2
    Assert.AreEqual 2, list.Count
    AssertOriginalList
End Sub

Public Sub TestRetrieveAfterAdding()
    list.Add 1
    Assert.AreEqual 1, list(0)
    AssertOriginalList
End Sub

Public Sub TestAddRangeFromCollection()
    Dim col As New Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    list.AddRange col
    Assert.AreEqual 3, list.Count
    Assert.AreEqual 1, list(0)
    Assert.AreEqual 2, list(1)
    Assert.AreEqual 3, list(2)
    AssertOriginalList
End Sub

Public Sub TestAddRangeFromArray()
    Dim a() As Long
    a = cArray.NewArray(ciLong, 1, 2, 3)
    
    list.AddRange a
    Assert.AreEqual 3, list.Count
    Assert.AreEqual 1, list(0)
    Assert.AreEqual 2, list(1)
    Assert.AreEqual 3, list(2)
    AssertOriginalList
End Sub

Public Sub TestSetItem()
    list.AddRange cArray.NewArray(ciLong, 1, 2, 3, 4, 5)
    list(2) = "Hello"
    Assert.AreEqual 5, list.Count
    Assert.AreEqual 1, list(0)
    Assert.AreEqual 2, list(1)
    Assert.AreEqual "Hello", list(2)
    Assert.AreEqual 4, list(3)
    Assert.AreEqual 5, list(4)
    AssertOriginalList
End Sub

Public Sub TestClear()
    list.Add 1
    list.Clear
    Assert.AreEqual 0, list.Count
    AssertOriginalList
End Sub

Public Sub TestCopyTo()
    Dim a(2) As String
    list.Add "Item1"
    list.Add "Item2"
    list.Add "Item3"
    list.CopyTo a
    Assert.AreEqual "Item1", a(0)
    Assert.AreEqual "Item2", a(1)
    Assert.AreEqual "Item3", a(2)
End Sub

Public Sub TestCopyToIndexOffset()
    Dim a(5) As String
    list.Add "Item1"
    list.Add "Item2"
    list.Add "Item3"
    list.CopyTo a, 2
    Assert.AreEqual "Item1", a(2)
    Assert.AreEqual "Item2", a(3)
    Assert.AreEqual "Item3", a(4)
End Sub

Public Sub TestEnumerator()
    list.AddRange cArray.NewArray(ciString, "Item1", "Item2", "Item3")
    
    Dim v As Variant
    For Each v In list
        Assert.AreEqual "Item", Left$(v, 4)
    Next v
End Sub

Public Sub TestEnumeratorIndex()
    Dim i As Long
    For i = 1 To 10
        list.Add i
    Next i
    
    Dim v As Variant
    i = 3
    For Each v In list.GetEnumerator(2, 5)
        Assert.AreEqual i, v
        i = i + 1
    Next v
End Sub

Public Sub TestIndexOf()
    Dim i As Long
    For i = 1 To 10
        list.Add "Item " & i
    Next i
    Assert.AreEqual 8, list.IndexOf("Item 9")
End Sub

Public Sub TestInsertAtBeginning()
    Dim i As Long
    For i = 1 To 10
        list.Add "Item " & i
    Next i
    list.Insert 0, 1
    Assert.AreEqual 11, list.Count
    Assert.AreEqual 1, list(0)
    For i = 1 To 10
        Assert.AreEqual "Item " & i, list(i)
    Next i
    AssertOriginalList
End Sub

Public Sub TestInsertAtEnd()
    Dim i As Long
    For i = 1 To 10
        list.Add "Item " & i
    Next i
    list.Insert 10, 1
    Assert.AreEqual 1, list(10)
    For i = 1 To 10
        Assert.AreEqual "Item " & i, list(i - 1)
    Next i
    AssertOriginalList
End Sub

Public Sub TestInsertRangeArray()
    list.AddRange cArray.NewArray(ciLong, 1, 2, 3, 4)
    list.Insertrange 2, cArray.NewArray(ciString, "Item1", "Item2", "Item3")
    AssertInsert
    AssertOriginalList
End Sub

Public Sub TestInsertRangecollection()
    list.AddRange cArray.NewArray(ciLong, 1, 2, 3, 4)
    Dim col As New Collection
    col.Add "Item1"
    col.Add "Item2"
    col.Add "Item3"
    list.Insertrange 2, col
    AssertInsert
    AssertOriginalList
End Sub

Public Sub TestLastIndexOf()
    Dim i As Long
    For i = 1 To 10
        list.Add "Item " & i
    Next i
    
    Assert.AreEqual 3, list.LastIndexOf("Item 4"), "LastIndexOf should return 3."
End Sub

Public Sub TestRemoveAtBeginning()
    Dim i As Long
    For i = 1 To 10
        list.Add "Item " & i
    Next i
        
    list.RemoveAt 0
    Assert.AreEqual 9, list.Count, "count should return 9."
    For i = 0 To 8
        Assert.AreEqual "Item " & i + 2, list(i), "list should return 'Item " & i + 2 & "'."
    Next i
    AssertOriginalList
End Sub

Public Sub TestRemoveAtEnd()
    Dim i As Long
    For i = 1 To 10
        list.Add "Item " & i
    Next i
    
    list.RemoveAt 9
    Assert.AreEqual 9, list.Count, "Count should return 9."
    For i = 0 To 8
        Assert.AreEqual "Item " & i + 1, list(i), "list should return 'Item " & i + 1 & "'."
    Next i
    AssertOriginalList
End Sub

Public Sub TestRemove()
    Dim i As Long
    For i = 1 To 10
        list.Add "Item " & i
    Next i
    
    list.Remove "Item 3"
    Assert.AreEqual 9, list.Count, "Count should return 9."
    Assert.AreEqual "Item 1", list(0), "list should return 'Item 1'."
    Assert.AreEqual "Item 2", list(1), "list should return 'Item 2'."
    For i = 2 To 8
        Assert.AreEqual "Item " & i + 2, list(i), "list should return 'Item " & i + 2 & "'."
    Next i
    AssertOriginalList
End Sub

Public Sub TestRemoveRange()
    Dim i As Long
    For i = 1 To 10
        list.Add "Item " & i
    Next i

    list.RemoveRange 2, 4
    Assert.AreEqual 6, list.Count, "Count should return 6."
    Assert.AreEqual "Item 1", list(0), "list should return 'Item 1'."
    Assert.AreEqual "Item 2", list(1), "list should return 'Item 2'."
    For i = 2 To 5
        Assert.AreEqual "Item " & i + 5, list(i), "list should return 'Item " & i + 5 & "'."
    Next i
    AssertOriginalList
End Sub

Public Sub TestReverse()
    Dim i As Long
    For i = 1 To 10
        list.Add i
    Next i
    list.Reverse
    For i = 0 To 9
        Assert.AreEqual 10 - i, list(i), "list should return " & 10 - i
    Next i
    AssertOriginalList
End Sub

Public Sub TestToArray()
    Dim i As Long
    For i = 1 To 10
        list.Add i
    Next i
    Dim a() As Long
    a = list.ToArray(ciLong)
    For i = 0 To 9
        Assert.AreEqual i + 1, a(i)
    Next i
End Sub

Public Sub TestSetRange()
    Dim i As Long
    For i = 1 To 10
        list.Add i
    Next i
    list.SetRange 2, cArray.NewArray(ciString, "Item 3", "Item 4", "Item 5")
    Assert.AreEqual 10, list.Count, "Count should return 10."
    For i = 0 To 9
        If i >= 2 And i <= 4 Then
            Assert.AreEqual "Item " & i + 1, list(i), "list should return 'Item " & i - 1 & "'."
        Else
            Assert.AreEqual i + 1, list(i)
        End If
    Next i
    AssertOriginalList
End Sub




Private Sub AssertInsert()
    Assert.AreEqual 7, list.Count, "Count should return 7."
    Assert.AreEqual 1, list(0), "list should return '1'."
    Assert.AreEqual 2, list(1), "list should return '2'."
    Assert.AreEqual "Item1", list(2), "list should return 'Item1'."
    Assert.AreEqual "Item2", list(3), "list should return 'Item2'."
    Assert.AreEqual "Item3", list(4), "list should return 'Item3'."
    Assert.AreEqual 3, list(5), "list should return '3'."
    Assert.AreEqual 4, list(6), "list should return '4'."
End Sub


Private Sub AssertOriginalList()
    Assert.AreEqual 20 + list.Count, original.Count, "Original count should be " & 20 + list.Count
    
    Dim i As Long
    For i = 0 To 3
        Assert.AreEqual i + 1, original(i), "original should return '" & i & "'."
    Next i
    For i = 4 + list.Count To 19 + list.Count
        Assert.AreEqual i - list.Count + 1, original(i), "original should return '" & i - list.Count + 1 & "'."
    Next i
End Sub


