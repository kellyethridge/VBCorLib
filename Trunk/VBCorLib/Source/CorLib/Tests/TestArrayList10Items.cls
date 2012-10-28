VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayList10Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private list As ArrayList
Private i As Long

Public Sub Setup()
    Set list = New ArrayList
    For i = 1 To 10
        list.Add i
    Next i
End Sub

Public Sub TestLastIndexOf()
    Assert.AreEqual 3, list.LastIndexOf(4&), "LastIndexOf should return 3."
End Sub

Public Sub TestRemoveAtBeginning()
    list.RemoveAt 0
    Assert.AreEqual 9, list.Count, "count should return 9."
    For i = 0 To 8
        Assert.AreEqual i + 2, list(i), "list should return '" & i + 2 & "'."
    Next i
End Sub

Public Sub TestRemoveAtEnd()
    list.RemoveAt 9
    Assert.AreEqual 9, list.Count, "Count should return 9."
    For i = 0 To 8
        Assert.AreEqual i + 1, list(i), "list should return '" & i + 1 & "'."
    Next i
End Sub

Public Sub TestRemove()
    list.Remove 3&
    Assert.AreEqual 9, list.Count, "Count should return 9."
    Assert.AreEqual 1, list(0), "list should return '1'."
    Assert.AreEqual 2, list(1), "list should return '2'."
    For i = 2 To 8
        Assert.AreEqual i + 2, list(i), "list should return '" & i + 2 & "'."
    Next i
End Sub

Public Sub TestRemoveRange()
    list.RemoveRange 2, 4
    Assert.AreEqual 6, list.Count, "Count should return 6."
    Assert.AreEqual 1, list(0), "list should return '1'."
    Assert.AreEqual 2, list(1), "list should return '2'."
    For i = 2 To 5
        Assert.AreEqual i + 5, list(i), "list should return '" & i + 5 & "'."
    Next i
End Sub

Public Sub TestReverse()
    list.Reverse
    For i = 0 To 9
        Assert.AreEqual 10 - i, list(i), "list should return " & 10 - i
    Next i
End Sub

Public Sub TestToArray()
    Dim a() As Long
    a = list.ToArray(ciLong)
    For i = 0 To 9
        Assert.AreEqual i + 1, a(i)
    Next i
End Sub

Public Sub TestSetRange()
    list.SetRange 2, cArray.NewArray(ciString, "Item 3", "Item 4", "Item 5")
    Assert.AreEqual 10, list.Count, "Count should return 10."
    For i = 0 To 9
        If i >= 2 And i <= 4 Then
            Assert.AreEqual "Item " & i + 1, list(i), "list should return 'Item " & i - 1 & "'."
        Else
            Assert.AreEqual i + 1, list(i)
        End If
    Next i
End Sub

Public Sub TestEnumeratorIndex()
    Dim v As Variant
    i = 3
    For Each v In list.GetEnumerator(2, 5)
        Assert.AreEqual i, v
        i = i + 1
    Next v
End Sub

Public Sub TestIndexOf()
    Assert.AreEqual 8, list.IndexOf(9&)
End Sub

Public Sub TestInsertAtBeginning()
    list.Insert 0, 1
    Assert.AreEqual 11, list.Count
    Assert.AreEqual 1, list(0)
    For i = 1 To 10
        Assert.AreEqual i, list(i)
    Next i
End Sub

Public Sub TestInsertAtEnd()
    list.Insert 10, 1
    Assert.AreEqual 1, list(10)
    For i = 0 To 9
        Assert.AreEqual i + 1, list(i)
    Next i
End Sub

Public Sub TestSerialize()
    Dim pp As New PropertyBag
    Dim pp2 As New PropertyBag
    Dim Copy As ArrayList
    
    list.Add Nothing
    list.Add Empty
    
    pp.WriteProperty "list", list
    pp2.Contents = pp.Contents
    Set Copy = pp2.ReadProperty("list")
    
    Dim i As Long
    For i = 0 To 9
        Assert.AreEqual i + 1, Copy(i)
    Next i
    
    Assert.That Copy(10), Iz.Nothing
    Assert.IsVBEmpty Copy(11)
End Sub
