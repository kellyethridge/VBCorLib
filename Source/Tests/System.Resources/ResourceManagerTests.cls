VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceManagerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ResourceManagerTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetString_WithSpecifiedCulture_ReturnsExpected"
            Test.Use NewCultureInfo("en-US"), "hello"
            Test.Use NewCultureInfo("or-IN"), "aloha"
    End Select
End Sub

Public Sub GetString_WithNoCultureInfo_ReturnsDefaultValue()
    Dim Manager As ResourceManager
    Set Manager = MakeManager
    
    Dim Actual As String
    Actual = Manager.GetString(101)
    
    Assert.That Actual, Iz.EqualTo("hello")
End Sub

Public Sub GetString_WithSpecifiedCulture_ReturnsExpected(ByVal Culture As CultureInfo, ByVal Expected As String)
    Dim Manager As ResourceManager
    Set Manager = MakeManager
    
    Dim Actual As String
    Actual = Manager.GetString(101, Culture)

    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetString_WithUnfoundResource_ReturnsEmptyString()
    Dim Manager As ResourceManager
    Set Manager = MakeManager
    
    Dim Actual As String
    Actual = Manager.GetString(102)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub GetObject_WithIconResource_ReturnsStdPicture()
     Dim Manager As ResourceManager
     Set Manager = MakeManager
     
     Dim Actual As IPictureDisp
     Set Actual = Manager.GetObject(1, ResourceTypes.IconResource)
     
     Assert.That Actual, Iz.Not.Nothing
     Assert.That Actual.Type, Iz.EqualTo(ResourceTypes.IconResource)
End Sub

Private Function MakeManager() As ResourceManager
    Set MakeManager = Cor.NewResourceManager(Path.Combine(App.Path, "CorLib.Tests.res"))
End Function
