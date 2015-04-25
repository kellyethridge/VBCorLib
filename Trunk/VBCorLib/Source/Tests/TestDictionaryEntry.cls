VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDictionaryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDictionaryEntry
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Collections"
End Sub

Public Sub TestEntry()
    Dim de As DictionaryEntry
    Set de = newdictionaryentry("key", "value")
    Assert.AreEqual "key", de.Key
    Assert.AreEqual "value", de.Value
End Sub

Public Sub TestSet()
    Dim de As DictionaryEntry
    Set de = New DictionaryEntry
    de.Key = "key"
    de.Value = "value"
    Assert.AreEqual "key", de.Key
    Assert.AreEqual "value", de.Value
End Sub
