VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSortedList"
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

Public Sub Setup()
    Set List = New SortedList
End Sub

Public Sub TestEmptyList()
    Assert.AreEqual 0, List.Count, "Count should be 0."
End Sub

Public Sub TestCountAfterAddingOne()
    List.Add 1, "hi"
    Assert.AreEqual 1, List.Count, "Count should be 1."
End Sub

Public Sub TestCountAfterAddingTwo()
    List.Add 1, "hi"
    List.Add 2, "Hello"
    Assert.AreEqual 2, List.Count, "Count should be 2."
End Sub

Public Sub TestAddingDuplicateKey()
    List.Add 1, "hi"
    On Error GoTo errTrap
    List.Add 1, "bye"
    On Error GoTo 0
    Assert.Fail "An error should be raised."
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "A exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestAddingDifferentKeyTypes()
    List.Add 1, "hi"
    On Error GoTo errTrap
    List.Add "1", "byte"
    On Error GoTo 0
    Assert.Fail "Error should be raised."
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "Exception should be caught."
    Assert.IsTrue TypeOf Ex Is InvalidOperationException, "Wrong exception type."
End Sub

Public Sub TestContents()
    List.Add "1", "hi"
    List.Add "2", "Hello"
    List.Add "key", "Me"

    Assert.AreEqual "Me", List("key"), "List should return 'Me'."
End Sub

Public Sub TestGetByIndex()
    List.Add 3, "Me"
    List.Add 1, "hi"
    List.Add 2, "Hello"

    Assert.AreEqual "hi", List.GetByIndex(0)
    Assert.AreEqual "Hello", List.GetByIndex(1)
    Assert.AreEqual "Me", List.GetByIndex(2)
End Sub

Public Sub TestClone()
    Dim i As Long
    For i = 1 To 10
        List.Add i, i
    Next i
    Dim list2 As SortedList
    Set list2 = List.Clone
    For i = 1 To 10
        Assert.AreEqual i, list2(i)
    Next i
End Sub

Public Sub TestContains()
    Dim i As Long
    For i = 1 To 10
        List.Add i, i
    Next i
    Assert.IsTrue List.Contains(3&)
End Sub

Public Sub TestContainsValue()
    Dim i As Long
    For i = 1 To 10
        List.Add i, "Item" & i
    Next i
    Assert.IsTrue List.ContainsValue("Item3")
End Sub

Public Sub TestCopyTo()
    Dim i As Long
    For i = 1 To 10
        List.Add i, "Item" & i
    Next i
    Dim a() As DictionaryEntry
    ReDim a(9)
    List.CopyTo a, 0
    For i = 0 To 9
        Assert.AreEqual i + 1, a(i).Key
        Assert.AreEqual "Item" & i + 1, a(i).Value
    Next i
End Sub

Public Sub TestGetKey()
    Dim i As Long
    For i = 1 To 10
        List.Add i, i
    Next i
    Assert.AreEqual 6, List.GetKey(5), "Should return 6."
End Sub

Public Sub TestIndexOfKey()
    Dim i As Long
    For i = 1 To 10
        List.Add i, i
    Next i
    Assert.AreEqual 5, List.IndexOfKey(6&), "Should return 5."
End Sub

Public Sub TestIndexOfValue()
    Dim i As Long
    For i = 1 To 10
        List.Add i, i
    Next i
    Assert.AreEqual 5, List.IndexOfValue(6&), "Should return 5."
End Sub

Public Sub TestRemove()
    Dim i As Long
    For i = 1 To 10
        List.Add i, i
    Next i
    List.Remove 5&
    Assert.AreEqual 9, List.Count, "Count should be 9."
    For i = 1 To 9
        Assert.AreEqual IIf(i < 5, i, i + 1), List.GetKey(i - 1)
    Next i
End Sub

Public Sub TestRemoveAt()
    Dim i As Long
    For i = 1 To 10
        List.Add i, i
    Next i
    List.RemoveAt 4
    Assert.AreEqual 9, List.Count, "Count should be 9."
    For i = 1 To 9
        Assert.AreEqual IIf(i < 5, i, i + 1), List.GetKey(i - 1)
    Next i
End Sub

Public Sub TestSetByIndex()
    Dim i As Long
    For i = 1 To 10
        List.Add i, i
    Next i
    List.SetByIndex 4, "Hello"
    Assert.AreEqual 10, List.Count, "Count should not change."
    For i = 0 To 9
        Assert.AreEqual IIf(i = 4, "Hello", i + 1), List.GetByIndex(i)
    Next i
End Sub

Public Sub TestTrimToSize()
    Dim i As Long
    For i = 1 To 5
        List.Add i, i
    Next i
    List.TrimToSize
    Assert.AreEqual 5, List.Capacity, "Capacity should be 5."
End Sub

Public Sub TestEnumerator()
    Dim i As Long
    For i = 1 To 10
        List.Add i, i
    Next i
    Dim Entry As DictionaryEntry
    i = 1
    For Each Entry In List
        Assert.AreEqual i, Entry.Key
        Assert.AreEqual i, Entry.Value
        i = i + 1
    Next Entry
End Sub

Public Sub TestCreateWithList()
    Dim l2 As SortedList
    Dim i As Long

    For i = 1 To 10
        List.Add i, i
    Next i
    Set l2 = newsortedlist(List)
    Assert.AreEqual 10, l2.Capacity, "Capacity should be 10."
    For i = 1 To 10
        Assert.AreEqual i, l2.GetKey(i - 1)
    Next i
End Sub

Public Sub TestPersitence()
    Dim l2 As SortedList
    Dim i As Long
    For i = 1 To 100
        List.Add i, i
    Next i
    Dim pp1 As New PropertyBag
    Dim pp2 As New PropertyBag
    pp1.WriteProperty "list", List
    pp2.Contents = pp1.Contents
    Set l2 = pp2.ReadProperty("list")

    Assert.AreEqual 100, l2.Count
    For i = 1 To 10
        Assert.AreEqual i, l2(i)
    Next i
End Sub

Public Sub TestValueAndKeyList()
    List.Add 1, "hi"
    Assert.AreEqual 1, List.GetKeyList(0)
    Assert.AreEqual "hi", List.GetValueList(0)
End Sub

Public Sub TestEnumerateKeysAndValues()
    List.Add 1, "value0"
    List.Add 2, "value1"
    
    Dim v As Variant
    Dim i As Long
    For Each v In List.Keys
        Assert.AreEqual List.GetKey(i), v
        i = i + 1
    Next v
    
    i = 0
    For Each v In List.Values
        Assert.AreEqual "value" & i, v
        i = i + 1
    Next v
End Sub


