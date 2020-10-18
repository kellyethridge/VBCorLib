VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceReaderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ResourceReaderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub GetEnumerator_WhenCalled_ReturnsEnumerableResources()
    Dim Resources As New Hashtable
    Dim Reader As ResourceReader
    Dim Enumerator As IDictionaryEnumerator
    
    Set Reader = Cor.NewResourceReader(Path.Combine(App.Path, "CorLib.Tests.RES"))
    Set Enumerator = Reader.GetEnumerator
    
    Do While Enumerator.MoveNext
        Resources.Add Enumerator.Key, Enumerator.Value
    Loop
    
    AssertResourceExists Resources, 101, StringResource, 1033
    AssertResourceExists Resources, 101, ResourceTypes.GroupIcon, 1033
    AssertResourceExists Resources, 1, ResourceTypes.IconResource, 1033
    AssertResourceExists Resources, 101, ResourceTypes.GroupCursor, 1033
    AssertResourceExists Resources, 1, ResourceTypes.CursorResource, 1033
    AssertResourceExists Resources, 101, "CUSTOM", 1033
    Assert.That Resources(Cor.NewResourceKey(101, ResourceTypes.StringResource, 1033)), Iz.EqualTo("hello")
End Sub

Private Sub AssertResourceExists(ByVal Resources As Hashtable, ByRef ResourceName As Variant, ByRef ResourceType As Variant, ByVal LanguageId As Long)
    Dim Key As ResourceKey
    
    Set Key = Cor.NewResourceKey(ResourceName, ResourceType, LanguageId)
    Assert.That Resources.ContainsKey(Key), Iz.True
End Sub


