VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceSetTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ResourceSetTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub GetString_WithenUSCulture_ReturnsString()
    Dim RS As ResourceSet
    Set RS = MakeResourceSet
    
    Dim Actual As String
    Actual = RS.GetString(101)
    
    Assert.That Actual, Iz.EqualTo("hello")
End Sub

Public Sub GetString_WhenNotExisting_ReturnsEmpty()
    Dim RS As ResourceSet
    Set RS = MakeResourceSet
    
    Dim Actual As String
    Actual = RS.GetString(999)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub GetObject_WhenIsByteArray_ReturnsByteArray()
    Dim RS As ResourceSet
    Set RS = MakeResourceSet
    
    Dim Actual() As Byte
    Actual = RS.GetObject(101, "Custom")
    
    Assert.That Actual, Iz.EqualTo(NewBytes(48, 49, 50, 51, 52, 53, 54, 55, 56, 57))
End Sub

Public Sub GetObject_WithNonExsiting_ReturnsEmpty()
    Dim RS As ResourceSet
    Set RS = MakeResourceSet
    
    Dim Actual As Variant
    Actual = RS.GetObject(101, "NotFound")
    
    Assert.That Actual, Iz.vbEmpty
End Sub

Public Sub Enumeration_WhenCalled_EnumeratedResources()
    Dim RS As ResourceSet
    Dim Entry As DictionaryEntry
    Dim Resources As New Hashtable
    Set RS = MakeResourceSet
    
    For Each Entry In RS
        Resources.Add Entry.Key, Entry.Value
    Next
    
    Assert.That Resources.Count, Iz.EqualTo(7)
    Assert.That Resources.ContainsKey(Cor.NewResourceKey(101, ResourceTypes.StringResource, 1033)), Iz.True
End Sub

Public Sub GetString_UsingDifferentCulture_ReturnsExpected()
    Dim RS As ResourceSet
    Set RS = MakeResourceSet("or-IN")
    
    Dim Actual As String
    Actual = RS.GetString(101)
    
    Assert.That Actual, Iz.EqualTo("aloha")
End Sub

Private Function MakeResourceSet(Optional ByVal CultureId As String = "en-US") As ResourceSet
    Set MakeResourceSet = Cor.NewResourceSet(Cor.NewCultureInfo(CultureId), Path.Combine(App.Path, "CorLib.Tests.RES"))
End Function
