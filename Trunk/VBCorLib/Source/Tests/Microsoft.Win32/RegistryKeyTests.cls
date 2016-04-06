VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKeyTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' RegistryKeyTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Const RegistryPathBase              As String = "software\vb and vba program settings"
Private Const RegistryKeysPath              As String = "software\vb and vba program settings\VBCorLibTests"
Private Const RegistryValuesPath            As String = "software\vb and vba program settings\VBCorLibTests\Values"
'Private Const RegistryKeyPath           As String = "HKEY_CURRENT_USER\" & RegistryPath
Private Const REG_OPTION_NON_VOLATILE   As Long = 0
Private Const REG_DWORD                 As Long = 4

Private Declare Function RegQueryValueExStr Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Private mKey As RegistryKey
Private hKey As Long

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "Microsoft.Win32"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "RoundTrip_WithValues_ReturnsExpected"
            Test.Use CLng(1)
            Test.Use "One"
            Test.Use CCur(1)
            Test.Use NewBytes(1, 2, 3, 4)
            Test.Use NewStrings("one", "two", "three")
            Test.Use "%PATH%"
        
        Case "GetValueKey_WhenCalled_ReturnsExpected"
            Test.Use CLng(1), RegistryValueKind.DWordKind
            Test.Use "One", RegistryValueKind.StringKind
            Test.Use CCur(1), RegistryValueKind.QWordKind
            Test.Use NewBytes(1, 2, 3, 4), RegistryValueKind.BinaryKind
            Test.Use NewStrings("one", "two", "three"), RegistryValueKind.MultiStringKind
            
    End Select
End Sub

Private Function Rooted(ByRef Path As String) As String
    Rooted = "HKEY_CURRENT_USER\" & Path
End Function

Public Sub Teardown()
    On Error Resume Next
    DeleteSetting "VBCorLibTests"
End Sub

Public Sub OpenSubKey_WithExistingPath_ReturnsKey()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim Actual As RegistryKey
    
    Set Actual = Registry.GetSubKey(Rooted(RegistryValuesPath))
    
    Assert.That Actual, Iz.Not.Nothing
    Assert.That Actual.Name, Iz.EqualTo(Rooted(RegistryValuesPath))
End Sub

Public Sub SubKeyCount_WhenCalled_ReturnsExpected()
    SaveSetting "VBCorLibTests", "Values1", "Key", "Value"
    SaveSetting "VBCorLibTests", "Values2", "Key", "Value"
    SaveSetting "VBCorLibTests", "Values3", "Key", "Value"
    Dim Key As RegistryKey
    Set Key = Registry.GetSubKey(Rooted(RegistryKeysPath))

    Dim Actual As Long
    Actual = Key.SubKeyCount

    Assert.That Actual, Iz.EqualTo(3)
End Sub

Public Sub ValueCount_WhenCalled_ReturnsExpected()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    SaveSetting "VBCorLibTests", "Values", "Key2", "Value2"
    SaveSetting "VBCorLibTests", "Values", "Key3", "Value3"
    Dim Key As RegistryKey
    Set Key = Registry.GetSubKey(Rooted(RegistryValuesPath))

    Dim Actual As Long
    Actual = Key.ValueCount

    Assert.That Actual, Iz.EqualTo(3)
End Sub

Public Sub CreateSubKey_WhenCalled_CreatesKey()
    Registry.CurrentUser.CreateSubKey RegistryValuesPath & "1"
    Registry.CurrentUser.CreateSubKey RegistryValuesPath & "2"
    
    Dim ParentKey As RegistryKey
    Set ParentKey = Registry.GetSubKey(Rooted(RegistryKeysPath))
    
    Assert.That ParentKey, Iz.Not.Nothing
    Assert.That ParentKey.SubKeyCount, Iz.EqualTo(2)
End Sub

Public Sub GetValue_WhenCalled_ReturnsValue()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryValuesPath)
    
    Dim Actual As Variant
    Actual = Key.GetValue("Key1")
    
    Assert.That Actual, Iz.EqualTo("Value1")
End Sub

Public Sub SetValue_WithValue_StoresValue()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryValuesPath, True)

    Key.SetValue "Key1", "NewValue"
    
    Dim Actual As String
    Actual = GetSetting("VBCorLibTests", "Values", "Key1")
    Assert.That Actual, Iz.EqualTo("NewValue")
End Sub

Public Sub SetValue_WithNonWritableKey_ThrowsUnauthroizedAccessException()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryValuesPath)

    On Error Resume Next
    Key.SetValue "Key1", "NewValue"
    
    AssertUnauthorizedAccessException Err
End Sub

Public Sub SetValue_WithCreatedSubKey_SetsValue()
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.CreateSubKey(RegistryKeysPath & "\MyValues")

    Key.SetValue "MyKey", "NewValue"
    
    Dim Actual As String
    Actual = GetSetting("VBCorLibTests", "MyValues", "MyKey")
    Assert.That Actual, Iz.EqualTo("NewValue")
End Sub

Public Sub CreateSubKey_WithReadOnlyKey_ThrowsUnauthorizedAccessException()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim ReadOnlyKey As RegistryKey
    Set ReadOnlyKey = Registry.CurrentUser.OpenSubKey(RegistryKeysPath)
    
    On Error Resume Next
    Dim Key As RegistryKey
    Set Key = ReadOnlyKey.CreateSubKey("ReadOnly")
    
    AssertUnauthorizedAccessException Err
End Sub

Public Sub OpenSubKey_WithNonExistingKey_ReturnsNothing()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim Actual As RegistryKey
    
    Set Actual = Registry.CurrentUser.OpenSubKey(RegistryKeysPath & "\Missing")
    
    Assert.That Actual, Iz.Nothing
End Sub

Public Sub DeleteSubKey_WhenCalled_RemovesSubKey()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryKeysPath, True)
    
    Key.DeleteSubKey "Values"
    
    Dim Actual As String
    Actual = GetSetting("VBCorLibTests", "Values", "Key1", "NewValue")
    Assert.That Actual, Iz.EqualTo("NewValue")
End Sub

Public Sub DeleteSubKey_WhenNonExisting_ThrowsArgumentException()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryKeysPath, True)

    On Error Resume Next
    Key.DeleteSubKey "Missing"
    
    AssertArgumentException Err
End Sub

Public Sub DeleteSubKey_WhenNonExistingAndNotThrowing_DoesNothing()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryKeysPath, True)

    Key.DeleteSubKey "Missing", False
End Sub

Public Sub DeleteSubKey_WhenContainsSubKeys_ThrowsInvalidOperationException()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryPathBase, True)
    
    On Error Resume Next
    Key.DeleteSubKey "VBCorLibTests"
    
    AssertInvalidOperationException Err
End Sub

Public Sub GetSubKeyNames_WhenCalled_ReturnsArrayOfNames()
    SaveSetting "VBCorLibTests", "Values1", "Key", "Value"
    SaveSetting "VBCorLibTests", "Values2", "Key", "Value"
    SaveSetting "VBCorLibTests", "Values3", "Key", "Value"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryKeysPath)
    
    Dim Actual() As String
    Actual = Key.GetSubKeyNames
    
    Dim Expected() As String
    Expected = NewStrings("Values1", "Values2", "Values3")
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetSubKeyNames_WithNoSubKeys_ReturnsEmptyArray()
    SaveSetting "VBCorLibTests", "Values", "Key", "Value"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryValuesPath)
    
    Dim Actual() As String
    Actual = Key.GetSubKeyNames
    
    Assert.That CorArray.IsNull(Actual), Iz.False
    Assert.That Actual, Iz.Empty
End Sub

Public Sub DeleteSubKeyTree_WithExistingKey_DeletesSubKeys()
    SaveSetting "VBCorLibTests", "Values1", "Key", "Value"
    SaveSetting "VBCorLibTests", "Values2", "Key", "Value"
    SaveSetting "VBCorLibTests", "Values3", "Key", "Value"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryPathBase, True)
    
    Key.DeleteSubKeyTree "VBCorLibTests"
    
    Dim Deleted As RegistryKey
    Set Deleted = Registry.CurrentUser.OpenSubKey(RegistryKeysPath)
    Assert.That Deleted, Iz.Nothing
End Sub

Public Sub CreateSubKey_WhenKeyExists_ReturnsOpenKey()
    SaveSetting "VBCorLibTests", "Values1", "Key", "Value"
    Dim Key As RegistryKey
    
    Set Key = Registry.CurrentUser.CreateSubKey(RegistryKeysPath & "\Values1")
    
    Assert.That Key, Iz.Not.Nothing
    Assert.That Key.GetValue("Key", ""), Iz.EqualTo("Value")
End Sub

Public Sub GetValueNames_WithNoValues_ReturnsEmptyArray()
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.CreateSubKey(RegistryValuesPath)
    
    Dim Actual() As String
    Actual = Key.GetValueNames
    
    Assert.That CorArray.IsNull(Actual), Iz.False
    Assert.That Actual, Iz.Empty
End Sub

Public Sub GetValueNames_WithValues_ReturnsNames()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value"
    SaveSetting "VBCorLibTests", "Values", "Key2", "Value"
    SaveSetting "VBCorLibTests", "Values", "Key3", "Value"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.OpenSubKey(RegistryValuesPath)
    
    Dim Actual() As String
    Actual = Key.GetValueNames
    
    Dim Expected() As String
    Expected = NewStrings("Key1", "Key2", "Key3")
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub RoundTrip_WithValues_ReturnsExpected(ByRef Value As Variant)
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.CreateSubKey(RegistryValuesPath)
    
    Key.SetValue "TestValue", Value
    Dim Actual As Variant
    Actual = Key.GetValue("TestValue")
    
    Assert.That Actual, Iz.EqualTo(Value).Strict
End Sub

Public Sub RoundTrip_WithExpandableValue_ReturnsExpanded()
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.CreateSubKey(RegistryValuesPath)
    
    Key.SetValue "TestValue", "%PATH%", ExpandStringKind
    Dim Actual As String
    Actual = Key.GetValue("TestValue")
    
    Assert.That Actual, Iz.EqualTo(Environ("PATH"))
End Sub

Public Sub DeleteValue_WhenCalled_RemovesValueFromRegistry()
    SaveSetting "VBCorLibTests", "Values", "Key1", "Value1"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.CreateSubKey(RegistryValuesPath)
    
    Key.DeleteValue "Key1"
    
    Dim Actual As String
    Actual = GetSetting("VBCorLibTests", "Values", "Key1", "")
    Assert.That Actual, Iz.Empty
End Sub

Public Sub DeleteValue_WithMissingValue_ThrowsArgumentException()
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.CreateSubKey(RegistryValuesPath)
    
    On Error Resume Next
    Key.DeleteValue "Key1"
    
    AssertArgumentException Err
End Sub

Public Sub DeleteSubKey_WhenCalled_RemovesSubKeyFromRegistry()
    SaveSetting "VBCorLibTests", "Values1", "Key1", "Value1"
    SaveSetting "VBCorLibTests", "Values2", "Key2", "Value2"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.CreateSubKey(RegistryKeysPath)

    Key.DeleteSubKey "Values1"
    
    Dim Actual() As String
    Actual = Key.GetSubKeyNames
    Assert.That Actual, Iz.EqualTo(Array("Values2"))
End Sub

Public Sub DeleteSubKeyTree_WhenCalled_RemovesSubKeysFromRegistry()
    Dim Root As RegistryKey
    Set Root = Registry.CurrentUser.CreateSubKey(RegistryKeysPath & "\Root")
    Root.CreateSubKey "Sub1"
    Root.CreateSubKey "Sub2"
    Root.CreateSubKey "Sub3"
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.CreateSubKey(RegistryKeysPath)

    Key.DeleteSubKeyTree "Root"
    
    Dim Actual() As String
    Actual = Key.GetSubKeyNames
    Assert.That Actual, Iz.Empty
End Sub

Public Sub GetValueKey_WhenCalled_ReturnsExpected(ByRef Value As Variant, ByVal Expected As RegistryValueKind)
    Dim Key As RegistryKey
    Set Key = Registry.CurrentUser.CreateSubKey(RegistryValuesPath)
    Key.SetValue "TestValue", Value
    
    Dim Actual As RegistryValueKind
    Actual = Key.GetValueKind("TestValue")
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub



''Public Sub TestRemote()
''    Dim mKey As RegistryKey
''    Dim s() As String
''    Dim i As Long
''
''    Set mKey = RegistryKey.OpenRemoteBaseKey(CurrentUser, "Soze")
''    Assert.IsNotNothing mKey
''    s = mKey.GetSubKeyNames
''    For i = 0 To UBound(s)
''        WriteLine s(i)
''    Next i
''    mKey.CloseKey
''End Sub
'
'
