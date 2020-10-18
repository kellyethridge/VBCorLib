VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' RegistryTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "Microsoft.Win32"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetRootKey_WithValue_ReturnsExpected"
            Test.Use "hkey_local_machine\somekey", Registry.LocalMachine
            Test.Use "hkey_current_user\somekey", Registry.CurrentUser
            Test.Use "hkey_current_config", Registry.CurrentConfig
            Test.Use "hkey_performance_data\somekey", Registry.PerformanceData
            Test.Use "hkey_users\users", Registry.Users
    End Select
End Sub

Public Sub Teardown()
    On Error Resume Next
    DeleteSetting "VBCorLibTests"
End Sub

Public Sub RootNames_WhenCalled_ReturnsExpected()
    Assert.That Registry.ClassesRoot.Name, Iz.EqualTo("HKEY_CLASSES_ROOT")
    Assert.That Registry.LocalMachine.Name, Iz.EqualTo("HKEY_LOCAL_MACHINE")
    Assert.That Registry.CurrentUser.Name, Iz.EqualTo("HKEY_CURRENT_USER")
    Assert.That Registry.CurrentConfig.Name, Iz.EqualTo("HKEY_CURRENT_CONFIG")
    Assert.That Registry.PerformanceData.Name, Iz.EqualTo("HKEY_PERFORMANCE_DATA")
    Assert.That Registry.Users.Name, Iz.EqualTo("HKEY_USERS")
End Sub

Public Sub SetValue_WhenCalled_StoresValueInRegistry()
    Registry.SetValue "HKEY_CURRENT_USER\Software\vb and vba program settings\VBCorLibTests\TestValues", "Key", "Value"
    
    Dim Actual As String
    Actual = GetSetting("VBCorLibTests", "TestValues", "Key", "")
    Assert.That Actual, Iz.EqualTo("Value")
End Sub

Public Sub GetValue_WhenCalled_ReturnsStoredValue()
    SaveSetting "VBCorLibTests", "TestValues", "Key", "Value"
    
    Dim Actual As String
    Actual = Registry.GetValue("HKEY_CURRENT_USER\Software\vb and vba program settings\VBCorLibTests\TestValues", "Key", "")
    
    Assert.That Actual, Iz.EqualTo("Value")
End Sub

Public Sub GetValue_WithInvalidBaseKeyName_ThrowsArgumentException()
    On Error Resume Next
    Registry.GetValue "HKEY_INVALID_NAME\Software\vb and vba program settings\VBCorLibTests\TestValues", "Key", ""
    
    AssertArgumentException Err
End Sub

Public Sub GetValue_WithEmptyKey_ThrowsArgumentException()
    On Error Resume Next
    Registry.GetValue "", "Key", ""
    
    AssertArgumentException Err
End Sub

Public Sub SetValue_WithInvalidBaseKeyName_ThrowsArgumentException()
    On Error Resume Next
    Registry.SetValue "HKEY_INVALID_NAME\Software\vb and vba program settings\VBCorLibTests\TestValues", "Key", "Value"
    
    AssertArgumentException Err
End Sub
