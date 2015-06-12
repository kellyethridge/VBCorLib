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
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Add_WithItemAndKey_AddsToTable"
            Test.Use 1, 1
            Test.Use "a", 1
            Test.Use "a", "A"
        
        Case "ContainsKey_WhenCalled_ReturnsExpected"
            Test.Use 1, 1, True
            Test.Use 1, 2, False
            Test.Use CLng(1), CInt(1), False
            Test.Use "1", 1, False
            Test.Use "1", "1", True
            Test.Use 1, "1", False
            
    End Select
End Sub

Public Sub Count_WithEmptyHashtable_ReturnsZero()
    Dim Table As New Hashtable
    
    Dim Actual As Long
    Actual = Table.Count
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub Add_WithItemAndKey_AddsToTable(ByRef Key As Variant, ByRef Value As Variant)
    Dim Table As New Hashtable
    
    Table.Add Key, Value
    
    Assert.That Table.Count, Iz.EqualTo(1)
    Assert.That Table(Key), Iz.EqualTo(Value)
End Sub

Public Sub Add_WithManyItems_AddsManyItems()
    Dim Table As New Hashtable
    
    Dim i As Long
    For i = 1 To 1000
        Table.Add "key" & i, i
    Next
    
    Assert.That Table.Count, Iz.EqualTo(1000)
End Sub

Public Sub ContainsKey_WhenCalled_ReturnsExpected(ByRef Key As Variant, ByRef SearchKey As Variant, ByVal Expected As Boolean)
    Dim Table As New Hashtable
    Table.Add Key, Key
    
    Dim Actual As Boolean
    Actual = Table.ContainsKey(SearchKey)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Item_WithNonExistingKey_AddsToTable()
    Dim Table As New Hashtable
    
    Table("abc") = 123
    
    Assert.That Table.ContainsKey("abc"), Iz.True
    Assert.That Table("abc"), Iz.EqualTo(123)
End Sub

Public Sub Item_WithExistingKey_ReplacesValue()
    Dim Table As New Hashtable
    Table.Add "abc", 123
    
    Table("abc") = 456
    
    Assert.That Table.ContainsKey("abc"), Iz.True
    Assert.That Table("abc"), Iz.EqualTo(456)
End Sub

Public Sub Remove_WithExistingKey_RemovesItem()
    Dim Table As New Hashtable
    Table.Add "abc", 123
    
    Table.Remove "abc"
    
    Assert.That Table.ContainsKey("abc"), Iz.False
    Assert.That Table.Count, Iz.EqualTo(0)
End Sub

Public Sub Remove_WithNonExistingKey_DoesNothing()
    Dim Table As New Hashtable
    Table.Add "abc", 123
    
    Table.Remove "def"
    
    Assert.That Table.ContainsKey("abc"), Iz.True
End Sub

Public Sub ContainsValue_WithNonExistingValue_ReturnsFalse()
    Dim Table As New Hashtable
    Table.Add "abc", 123
    
    Dim Actual As Boolean
    Actual = Table.ContainsValue(456)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub ContainsValue_WithExistingValue_ReturnsTrue()
    Dim Table As New Hashtable
    Table.Add "abc", 123
    
    Dim Actual As Boolean
    Actual = Table.ContainsValue(123)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Clone_WhenCalled_ReturnsCloneOfTable()
    Dim Table As New Hashtable
    Table.Add "abc", 123
    
    Dim Clone As Hashtable
    Set Clone = Table.Clone
    
    Assert.That Clone("abc"), Iz.EqualTo(123)
End Sub

Public Sub Clear_WhenCalled_RemovesAllItems()
    Dim Table As New Hashtable
    Table.Add "abc", 123
    
    Table.Clear
    
    Assert.That Table.Count, Iz.EqualTo(0)
End Sub

Public Sub Serialization_WhenCalled_SerializesInstance()
    Dim Table As New Hashtable
    Table.Add "abc", 123
    Table.Add "Nothing", Nothing
    Table.Add "Empty", Empty
    
    Dim Deserialized As Hashtable
    Set Deserialized = Persist(Table)
    
    Assert.That Deserialized.ContainsKey("abc"), Iz.True
    Assert.That Deserialized("abc"), Iz.EqualTo(123)
    Assert.That Deserialized("Nothing"), Iz.Nothing
    Assert.That Deserialized("Empty"), Iz.vbEmpty
End Sub

Public Sub CopyTo_WhenCalled_CopiesEntriesToArray()
    Dim Table As New Hashtable
    Table.Add "a", 1
    Table.Add "b", 2
    Dim Target(0 To 1) As Variant
    
    Table.CopyTo Target, 0
    
    Dim Expected As Variant
    Expected = Array(Cor.NewDictionaryEntry("a", 1), Cor.NewDictionaryEntry("b", 2))
    Assert.That Target, Iz.EquivalentTo(Expected).Using(New CorObjectComparer)
End Sub

Public Sub CopyTo_WithIndex_CopiesEntriesToArrayOffset()
    Dim Table As New Hashtable
    Table.Add "a", 1
    Table.Add "b", 2
    Dim Target(0 To 2) As Variant
    
    Table.CopyTo Target, 1
    
    Assert.That Target(0), Iz.vbEmpty
    Dim Expected As Variant
    Expected = Array(Empty, Cor.NewDictionaryEntry("a", 1), Cor.NewDictionaryEntry("b", 2))
    Assert.That Target, Iz.EquivalentTo(Expected).Using(New CorObjectComparer)
End Sub

Public Sub CopyTo_WithNullArray_ThrowsArgumentNullException()
    Dim Table As New Hashtable
    Dim NullTarget() As Variant
    
    On Error Resume Next
    Table.CopyTo NullTarget, 0
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub CopyTo_WithMultiDimArray_ThrowsArgumentException()
    Dim Table As New Hashtable
    Dim MultiDimTarget(0, 0) As Variant
    
    On Error Resume Next
    Table.CopyTo MultiDimTarget, 0
    
    AssertArgumentException Err, "Arr"
End Sub

Public Sub CopyTo_WithArrayIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Table As New Hashtable
    Table.Add "a", 1
    Dim Target(1 To 2) As Variant
    
    On Error Resume Next
    Table.CopyTo Target, 0
    
    AssertArgumentOutOfRangeException Err, "ArrayIndex"
End Sub

Public Sub CopyTo_WithIndexAndListCountGreaterThanUBound_ThrowsArgumentException()
    On Error Resume Next
    Dim List As New Hashtable
    List.Add 1, "a"
    List.Add 2, "b"
    Dim Target(0 To 1) As Variant
    
    List.CopyTo Target, 1
    
    AssertArgumentException Err
End Sub


