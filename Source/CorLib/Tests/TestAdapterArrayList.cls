VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAdapterArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestAdapterArrayList
'
Option Explicit

Private list As ArrayList

Public Sub Setup()
    Set list = ArrayList.Adapter(New ArrayList)
End Sub

Public Sub TestEmptyList()
    Assert.AreEqual 0, list.Count
End Sub

Public Sub TestAfterAddingOne()
    list.Add 1
    Assert.AreEqual 1, list.Count
End Sub

Public Sub TestAfterAddingTwo()
    list.Add 1
    list.Add 2
    Assert.AreEqual 2, list.Count
End Sub

Public Sub TestRetrieveAfterAdding()
    list.Add 1
    Assert.AreEqual 1, list(0)
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
End Sub

Public Sub TestAddRangeFromArray()
    Dim a() As Long
    a = cArray.NewArray(ciLong, 1, 2, 3)
    
    list.AddRange a
    Assert.AreEqual 3, list.Count
    Assert.AreEqual 1, list(0)
    Assert.AreEqual 2, list(1)
    Assert.AreEqual 3, list(2)
End Sub

Public Sub TestAddRangeFromICollection()
    Dim a As New ArrayList
    a.AddRange cArray.NewArray(ciLong, 1, 2, 3)
    
    list.AddRange a
    Assert.AreEqual 3, list.Count
    Assert.AreEqual 1, list(0)
    Assert.AreEqual 2, list(1)
    Assert.AreEqual 3, list(2)
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
End Sub

Public Sub TestClear()
    list.Add 1
    list.Clear
    Assert.AreEqual 0, list.Count
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
    list.AddRange cor.NewStrings("Item1", "Item2", "Item3")
    
    Dim i As Long
    i = 1
    
    Dim v As Variant
    For Each v In list
        Assert.AreEqual "Item" & i, v
        i = i + 1
    Next v
End Sub

Public Sub TestInsertRangeArray()
    list.AddRange cArray.NewArray(ciLong, 1, 2, 3, 4)
    list.InsertRange 2, cArray.NewArray(ciString, "Item1", "Item2", "Item3")
    AssertInsert
End Sub

Public Sub TestInsertRangecollection()
    list.AddRange cArray.NewArray(ciLong, 1, 2, 3, 4)
    Dim col As New Collection
    col.Add "Item1"
    col.Add "Item2"
    col.Add "Item3"
    list.InsertRange 2, col
    AssertInsert
End Sub

Public Sub TestContains()
    list.AddRange cArray.NewArray(ciLong, 1, 2, 3)
    Assert.IsTrue list.Contains(1&), "list should contain 1."
End Sub

Public Sub TestNotContains()
    list.AddRange cArray.NewArray(ciLong, 1, 2, 3)
    Assert.IsFalse list.Contains(4), "list should not contain 4."
End Sub

Public Sub TestToArrayWithEmptyList()
    Dim l() As Long
    l = list.ToArray(ciLong)
    Assert.AreEqual 0, cArray.GetLength(l)
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



