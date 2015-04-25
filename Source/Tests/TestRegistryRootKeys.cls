VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRegistryRootKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRegistryRootKeys
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "Microsoft.Win32"
End Sub

Public Sub Setup()
    DeleteTestValues
End Sub

Public Sub Teardown()
    DeleteTestValues
End Sub

Public Sub TestClassesRoot()
    Assert.AreEqual "HKEY_CLASSES_ROOT", Registry.ClassesRoot.Name
End Sub

Public Sub TestLocalMachine()
    Assert.AreEqual "HKEY_LOCAL_MACHINE", Registry.LocalMachine.Name
End Sub

Public Sub TestDynDataOn9x()
    If Environment.OSVersion.Platform = Win32NT Then Exit Sub
    Assert.AreEqual "HKEY_DYN_DATA", Registry.DynData.Name
End Sub

Public Sub TestDynDataOnNT()
    If Not Environment.OSVersion.Platform = Win32NT Then Exit Sub
    
    On Error GoTo errTrap
    Assert.AreEqual "HKEY_DYN_DATA", Registry.DynData.Name
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is IOException, "Wrong exception type caught."
End Sub

Public Sub TestSubKeyCount()
    Assert.IsTrue Registry.CurrentUser.SubKeyCount > 0, "Should be more than 0 sub keys."
End Sub

Public Sub TestSetGetValue()
    Registry.SetValue "HKEY_CURRENT_USER\Software\vb and vba program settings\VBCorLibTests\TestValues", "SetGet", "hello"
    Assert.AreEqual "hello", Registry.GetValue("HKEY_CURRENT_USER\Software\vb and vba program settings\VBCorLibTests\TestValues", "SetGet", "Not Found")
End Sub

Public Sub TestSetGetDefault()
    Registry.SetValue "HKEY_CURRENT_USER\software\vb and vba program settings\VBCorLibTests\TestValues", "", "Default"
    Assert.AreEqual "Default", Registry.GetValue("HKEY_CURRENT_USER\software\vb and vba program settings\VBCorLibTests\TestValues", "", "Not Found")
End Sub

Public Sub TestGetRootByName()
    Assert.AreEqual Registry.LocalMachine, Registry.GetRootKey("hkey_local_machine\somekey")
    Assert.AreEqual Registry.CurrentUser, Registry.GetRootKey("hkey_current_user")
End Sub

Private Sub DeleteTestValues()
    On Error Resume Next
    DeleteSetting "VBCorLibTests"
End Sub
