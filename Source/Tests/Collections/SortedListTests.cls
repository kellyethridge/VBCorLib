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
Option Explicit
Implements ICategorizable

Private List As SortedList

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
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

'Public Sub TestClone()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, i
'    Next i
'    Dim list2 As SortedList
'    Set list2 = List.Clone
'    For i = 1 To 10
'        Assert.AreEqual i, list2(i)
'    Next i
'End Sub
'
'Public Sub TestContains()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, i
'    Next i
'    Assert.IsTrue List.Contains(3&)
'End Sub
'
'Public Sub TestContainsValue()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, "Item" & i
'    Next i
'    Assert.IsTrue List.ContainsValue("Item3")
'End Sub
'
'Public Sub TestCopyTo()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, "Item" & i
'    Next i
'    Dim a() As DictionaryEntry
'    ReDim a(9)
'    List.CopyTo a, 0
'    For i = 0 To 9
'        Assert.AreEqual i + 1, a(i).Key
'        Assert.AreEqual "Item" & i + 1, a(i).Value
'    Next i
'End Sub
'
'Public Sub TestGetKey()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, i
'    Next i
'    Assert.AreEqual 6, List.GetKey(5), "Should return 6."
'End Sub
'
'Public Sub TestIndexOfKey()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, i
'    Next i
'    Assert.AreEqual 5, List.IndexOfKey(6&), "Should return 5."
'End Sub
'
'Public Sub TestIndexOfValue()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, i
'    Next i
'    Assert.AreEqual 5, List.IndexOfValue(6&), "Should return 5."
'End Sub
'
'Public Sub TestRemove()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, i
'    Next i
'    List.Remove 5&
'    Assert.AreEqual 9, List.Count, "Count should be 9."
'    For i = 1 To 9
'        Assert.AreEqual IIf(i < 5, i, i + 1), List.GetKey(i - 1)
'    Next i
'End Sub
'
'Public Sub TestRemoveAt()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, i
'    Next i
'    List.RemoveAt 4
'    Assert.AreEqual 9, List.Count, "Count should be 9."
'    For i = 1 To 9
'        Assert.AreEqual IIf(i < 5, i, i + 1), List.GetKey(i - 1)
'    Next i
'End Sub
'
'Public Sub TestSetByIndex()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, i
'    Next i
'    List.SetByIndex 4, "Hello"
'    Assert.AreEqual 10, List.Count, "Count should not change."
'    For i = 0 To 9
'        Assert.AreEqual IIf(i = 4, "Hello", i + 1), List.GetByIndex(i)
'    Next i
'End Sub
'
'Public Sub TestTrimToSize()
'    Dim i As Long
'    For i = 1 To 5
'        List.Add i, i
'    Next i
'    List.TrimToSize
'    Assert.AreEqual 5, List.Capacity, "Capacity should be 5."
'End Sub
'
'Public Sub TestEnumerator()
'    Dim i As Long
'    For i = 1 To 10
'        List.Add i, i
'    Next i
'    Dim Entry As DictionaryEntry
'    i = 1
'    For Each Entry In List
'        Assert.AreEqual i, Entry.Key
'        Assert.AreEqual i, Entry.Value
'        i = i + 1
'    Next Entry
'End Sub
'
'Public Sub TestCreateWithList()
'    Dim l2 As SortedList
'    Dim i As Long
'
'    For i = 1 To 10
'        List.Add i, i
'    Next i
'    Set l2 = newsortedlist(List)
'    Assert.AreEqual 10, l2.Capacity, "Capacity should be 10."
'    For i = 1 To 10
'        Assert.AreEqual i, l2.GetKey(i - 1)
'    Next i
'End Sub
'
'Public Sub TestPersitence()
'    Dim l2 As SortedList
'    Dim i As Long
'    For i = 1 To 100
'        List.Add i, i
'    Next i
'    Dim pp1 As New PropertyBag
'    Dim pp2 As New PropertyBag
'    pp1.WriteProperty "list", List
'    pp2.Contents = pp1.Contents
'    Set l2 = pp2.ReadProperty("list")
'
'    Assert.AreEqual 100, l2.Count
'    For i = 1 To 10
'        Assert.AreEqual i, l2(i)
'    Next i
'End Sub
'
'Public Sub TestValueAndKeyList()
'    List.Add 1, "hi"
'    Assert.AreEqual 1, List.GetKeyList(0)
'    Assert.AreEqual "hi", List.GetValueList(0)
'End Sub
'
'Public Sub TestEnumerateKeysAndValues()
'    List.Add 1, "value0"
'    List.Add 2, "value1"
'
'    Dim v As Variant
'    Dim i As Long
'    For Each v In List.Keys
'        Assert.AreEqual List.GetKey(i), v
'        i = i + 1
'    Next v
'
'    i = 0
'    For Each v In List.Values
'        Assert.AreEqual "value" & i, v
'        i = i + 1
'    Next v
'End Sub


