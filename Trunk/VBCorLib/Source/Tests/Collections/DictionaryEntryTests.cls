VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DictionaryEntryTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Equals_WithValues_ReturnsExpected"
            Test.Use 1, 1, False
            Test.Use "a", 2, False
            Test.Use "b", 1, False
            Test.Use "a", 1, True
            
    End Select
End Sub

Public Sub NewDictionaryEntry_WhenCalled_StoresValues()
    Dim Actual As DictionaryEntry
    
    Set Actual = Cor.NewDictionaryEntry("key", "value")
    
    Assert.That Actual.Key, Iz.EqualTo("key")
    Assert.That Actual.Value, Iz.EqualTo("value")
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByRef Key As Variant, ByRef Value As Variant, ByVal Expected As Boolean)
    Dim Entry1 As DictionaryEntry
    Dim Entry2 As DictionaryEntry
    Set Entry1 = Cor.NewDictionaryEntry("a", 1)
    Set Entry2 = Cor.NewDictionaryEntry(Key, Value)
    
    Dim Actual As Boolean
    Actual = Entry1.Equals(Entry2)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetHashCode_WithSameValues_ReturnsSame()
    Dim Entry1 As DictionaryEntry
    Dim Entry2 As DictionaryEntry
    Set Entry1 = Cor.NewDictionaryEntry("a", 1)
    Set Entry2 = Cor.NewDictionaryEntry("a", 1)
    
    Dim Actual1 As Long
    Dim Actual2 As Long
    Actual1 = Entry1.GetHashCode
    Actual2 = Entry2.GetHashCode
    
    Assert.That Actual1, Iz.EqualTo(Actual2)
End Sub

Public Sub Serialization_WhenCalled_Serializes()
    Dim Entry As DictionaryEntry
    Set Entry = Cor.NewDictionaryEntry("a", 1)
    
    Dim Actual As DictionaryEntry
    Set Actual = Persist(Entry)
    
    Assert.That Actual, Equals(Entry)
End Sub
