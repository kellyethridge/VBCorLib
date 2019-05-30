VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceKeyTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ResourceKeyTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Equals_WithNumericResourceNameAndType_ReturnsExpected"
            Test.Use 1, 2, 3, True
            Test.Use 1&, 2&, 3&, True
            Test.Use 1&, 2, 3, True
            Test.Use 1, 2&, 3, True
            Test.Use 1, 2, 0, True
            Test.Use 0, 2, 3, False
            Test.Use 1, 0, 3, False
            Test.Use 1, 2, -1, False
            Test.Use "1", 2, 3, False
            Test.Use 1, "2", 3, False
            
        Case "Equals_WithStringResourceNameAndType_ReturnsExpected"
            Test.Use "Name", "2", 3, True
            Test.Use "Name", "2", 0, True
            Test.Use "NAME", "2", 3, True
            Test.Use "NAME", "2", 0, True
            Test.Use "Name", 2, 3, False
            Test.Use "1", "2", 3, False
            Test.Use 1, "2", 3, False
        
    End Select
End Sub

Public Sub New_WithUnsupportedResourceType_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewResourceKey "Name", 1.1, 1
    
    AssertArgumentException Err
End Sub

Public Sub Equals_WithNumericResourceNameAndType_ReturnsExpected(ByVal Name As Variant, ByVal ResourceType As Variant, ByVal LanguageId As Long, ByVal Expected As Boolean)
    Dim Key     As ResourceKey
    Dim Other   As ResourceKey
    
    Set Other = Cor.NewResourceKey(1, 2, 3)
    Set Key = Cor.NewResourceKey(Name, ResourceType, LanguageId)
    
    Dim Actual As Boolean
    Actual = Key.Equals(Other)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithStringResourceNameAndType_ReturnsExpected(ByVal Name As Variant, ByVal ResourceType As Variant, ByVal LanguageId As Long, ByVal Expected As Boolean)
    Dim Key     As ResourceKey
    Dim Other   As ResourceKey
    
    Set Other = Cor.NewResourceKey("Name", "2", 3)
    Set Key = Cor.NewResourceKey(Name, ResourceType, LanguageId)
    
    Dim Actual As Boolean
    Actual = Key.Equals(Other)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub New_WithInvalidResourceName_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewResourceKey 1.1, "CUSTOM", 0
    
    AssertArgumentException Err, "ResourceName"
End Sub

Public Sub New_WithInvalidResourceType_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewResourceKey "CUSTOM", 1.1, 0
    
    AssertArgumentException Err, "ResourceType"
End Sub


