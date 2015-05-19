VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashtableTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHashTable
'
Option Explicit
Implements ICategorizable

Private t As Hashtable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Public Sub Count_WithEmptyHashtable_ReturnsZero()
    Dim Table As New Hashtable
    
    Dim Actual As Long
    Actual = Table.Count
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub Add_WithItemAndKey_AddsToTable()
    Dim Table As New Hashtable
    
    Table.Add "a", "a"
    
    Assert.That Table.Count, Iz.EqualTo(1)
    Assert.That Table("a"), Iz.EqualTo("a")
End Sub

Public Sub Add_WithManyItems_AddsManyItems()
    Dim Table As New Hashtable
    
    Dim i As Long
    For i = 1 To 1000
        Table.Add "key" & i, i
    Next
    
    Assert.That Table.Count, Iz.EqualTo(1000)
End Sub

'Public Sub TestContentsAfterAddingTwo()
'    t.Add 1, 1
'    t.Add 2, 2
'    Assert.AreEqual 1, t(1)
'    Assert.AreEqual 2, t(2)
'End Sub
'
'Public Sub Test1000Items()
'    Add1000Strings
'
'    Dim i As Long
'    For i = 1 To 1000
'        Assert.AreEqual "Value " & i, t("Key " & i)
'    Next i
'End Sub
'
'Public Sub TestChangeExisting()
'    t.Add "Key", "first"
'    Assert.AreEqual "first", t("Key")
'    t("Key") = "second"
'    Assert.AreEqual "second", t("Key")
'End Sub
'
'Public Sub TestSetsNew()
'    Assert.IsVBEmpty t("Key")
'    t("Key") = "first"
'    Assert.AreEqual "first", t("Key")
'End Sub
'
'Public Sub TestContainsKey()
'    Assert.IsFalse t.ContainsKey("Key")
'    t("Key") = "first"
'    Assert.IsTrue t.ContainsKey("Key")
'End Sub
'
'Public Sub TestRemove()
'    Add1000Longs
'
'    Dim i As Long
'    For i = 20 To 60
'        t.Remove i
'    Next i
'
'    For i = 1 To 1000
'        If i >= 20 And i <= 60 Then
'            Assert.IsVBEmpty t(i)
'        Else
'            Assert.AreEqual i, t(i)
'        End If
'    Next i
'End Sub
'
'Public Sub TestContainsValue()
'    Add1000Strings
'
'    Dim i As Long
'    Dim sw As StopWatch
'    Set sw = New StopWatch
'    sw.Reset
'    For i = 1 To 1000
'        Assert.IsTrue t.ContainsValue("Value " & i)
'    Next i
'    Assert.IsFalse t.ContainsValue("Not here.")
'End Sub
'
'Public Sub TestClone()
'    Add1000Strings
'
'    Dim i As Long
'    Dim t2 As Hashtable
'    Set t2 = t.Clone
'    For i = 1 To 1000
'        Assert.AreEqual "Value " & i, t2("Key " & i)
'    Next i
'End Sub
'
'Public Sub TestEnumerator()
'    Add1000Longs
'    Dim de As DictionaryEntry
'    Dim Col As New Collection
'    For Each de In t.GetEnumerator
'        Col.Add de.Value, CStr(de.Key)
'    Next de
'    Assert.AreEqual 1000, Col.Count
'End Sub
'
'Public Sub TestKeysCollection()
'    Add1000Longs
'    Assert.AreEqual 1000, t.Keys.Count
'    Dim v As Variant
'    Dim Col As New Collection
'    For Each v In t.Keys
'        Col.Add v, CStr(v)
'    Next v
'    Assert.AreEqual 1000, Col.Count
'End Sub
'
'Public Sub TestValuesCollection()
'    Add1000Longs
'    Assert.AreEqual 1000, t.Values.Count
'    Dim v As Variant
'    Dim Col As New Collection
'    For Each v In t.Values
'        Col.Add v, CStr(v)
'    Next v
'    Assert.AreEqual 1000, Col.Count
'End Sub
'
'Public Sub TestClear()
'    Add1000Longs
'    t.Clear
'    Assert.AreEqual 0, t.Count, "Count should return 0."
'End Sub
'
'Public Sub TestSerialization()
'    Dim h1 As New Hashtable
'    h1.Add "name", "kelly"
'    h1.Add "age", 22
'
'    Dim h2 As Hashtable
'    Dim p1 As New PropertyBag
'    Dim p2 As New PropertyBag
'
'    p1.WriteProperty "h", h1
'    p2.Contents = p1.Contents
'    Set h2 = p2.ReadProperty("h", Nothing)
'    Assert.IsNotNothing h2
'    Assert.AreEqual h1.Count, h2.Count
'    Assert.AreEqual h1("name"), h2("name")
'    Assert.AreEqual h1("age"), h2("age")
'End Sub
'
'Public Sub TestPersistValueOfNothingAndEmpty()
'    Dim h1 As New Hashtable
'    h1.Add "object", Nothing
'    h1.Add "variant", Empty
'
'    Dim h2 As Hashtable
'    Dim p1 As New PropertyBag
'    p1.WriteProperty "ht", h1
'
'    Set h2 = p1.ReadProperty("ht")
'
'    Assert.IsNothing h2("object")
'    Assert.IsVBEmpty h2("variant")
'
'End Sub
'
'
'
'Private Sub Add1000Longs()
'    Dim i As Long
'    For i = 1 To 1000
'        t.Add i, i
'    Next i
'End Sub
'Private Sub Add1000Strings()
'    Dim i As Long
'    For i = 1 To 1000
'        t.Add "Key " & i, "Value " & i
'    Next i
'End Sub
