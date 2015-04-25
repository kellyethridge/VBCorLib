VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRegistryKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRegistryKey
'
Option Explicit
Implements ICategorizable

Private Const REG_OPTION_NON_VOLATILE As Long = 0
Private Const REG_DWORD As Long = 4

Private Declare Function RegQueryValueExStr Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Private rk As RegistryKey
Private hKey As Long

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "Microsoft.Win32"
End Sub

Public Sub Setup()
    Set rk = Registry.CurrentUser.CreateSubKey("software\vb and vba program settings\VBCorLibTests\TestRegistryKey")
    RegOpenKeyExA &H80000001, "software\vb and vba program settings\VBCorLibTests\TestRegistryKey", 0, 983103, hKey
End Sub

Public Sub Teardown()
    RegCloseKey hKey
    rk.CloseKey
    Registry.CurrentUser.DeleteSubKeyTree "software\vb and vba program settings\VBCorLibTests"
End Sub

Public Sub TestCreateSubKey()
    Assert.IsNotNothing rk
    Assert.AreEqual "HKEY_CURRENT_USER\software\vb and vba program settings\VBCorLibTests\TestRegistryKey", rk.Name
End Sub

Public Sub TestSubKeyCount()
    Dim i As Long
    Dim j As Long
    For i = 1 To 5
        RegCreateKeyExA hKey, "SubKey " & i, 0, vbNullString, REG_OPTION_NON_VOLATILE, &HF003F, 0, j, 0
        RegCloseKey j
    Next i
    
    Assert.AreEqual 5, rk.SubKeyCount
End Sub

Public Sub TestValueCount()
    Dim i As Long
    For i = 1 To 10
        RegSetValueExA hKey, "Value " & i, 0, REG_DWORD, VarPtr(i), 4
    Next i
    
    Assert.AreEqual 10, rk.ValueCount
End Sub

Public Sub TestOpenSubKey()
    Dim NewKey As RegistryKey
    Set NewKey = Registry.CurrentUser.OpenSubKey("software\vb and vba program settings\VBCorLibTests\TestRegistryKey")

    Assert.IsNotNothing rk
    Assert.AreEqual "HKEY_CURRENT_USER\software\vb and vba program settings\VBCorLibTests\TestRegistryKey", rk.Name
    NewKey.CloseKey
End Sub

Public Sub TestCreateSubKeyWritable()
    Dim NewKey As RegistryKey
    Set NewKey = rk.CreateSubKey("WritableSubKey")
    
    Assert.IsNotNothing rk
    Assert.AreEqual "HKEY_CURRENT_USER\software\vb and vba program settings\VBCorLibTests\TestRegistryKey\WritableSubKey", NewKey.Name
    
    Dim hk As Long
    Assert.AreEqual 0, RegOpenKeyExA(hKey, "WritableSubKey", 0, &HF003F, hk)
    Assert.That 0, Iz.Not.EqualTo(hk)
    RegCloseKey hk
    
    NewKey.CloseKey
End Sub

Public Sub TestCreateSubKeyNonWritable()
    Dim hk As Long
    RegCreateKeyExA hKey, "SubKey", 0, vbNullString, REG_OPTION_NON_VOLATILE, &HF003F, 0, hk, 0
    RegCloseKey hk
    
    Dim NewKey As RegistryKey
    Set NewKey = rk.OpenSubKey("SubKey")
    
    On Error GoTo errTrap
    Dim CantWrite As RegistryKey
    Set CantWrite = NewKey.CreateSubKey("NotWritable")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."

errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is UnauthorizedAccessException, "Wrone exception type caught."
End Sub

Public Sub TestSubKeyNotFound()
    Dim NewKey As RegistryKey
    Set NewKey = rk.OpenSubKey("Missing")
    Assert.IsNothing NewKey
End Sub

Public Sub TestDeleteSubKey()
    Dim hk As Long
    RegCreateKeyExA hKey, "DeleteMe", 0, vbNullString, REG_OPTION_NON_VOLATILE, &HF003F, 0, hk, 0
    RegCloseKey hk
    
    rk.DeleteSubKey "DeleteMe"
    Assert.IsNothing rk.OpenSubKey("DeleteMe")
End Sub

Public Sub TestDeleteMissingSubKeyThrowable()
    On Error GoTo errTrap
    rk.DeleteSubKey "Missing"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."

errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestDeleteSubKeyWithSubKeysThrowable()
    Dim hk As Long
    RegCreateKeyExA hKey, "SubKey\ChildSubKey", 0, vbNullString, REG_OPTION_NON_VOLATILE, &HF003F, 0, hk, 0
    RegCloseKey hk
    
    On Error GoTo errTrap
    rk.DeleteSubKey "SubKey"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."

errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "an exception should be caught."
    Assert.IsTrue TypeOf Ex Is InvalidOperationException, "Wrong exception type caught."
End Sub
'
Public Sub TestGetSubKeyNames()
    Dim i As Long
    Dim j As Long
    For i = 1 To 5
        RegCreateKeyExA hKey, "SubKey " & i, 0, vbNullString, REG_OPTION_NON_VOLATILE, &HF003F, 0, j, 0
        RegCloseKey j
    Next i
    
    Dim s() As String
    s = rk.GetSubKeyNames
    
    Assert.AreEqual 5, cArray.GetLength(s), "Should contain 3 elements."
    
    For i = 1 To 5
'        Assert.Contains "SubKey " & i, s
        Assert.That s, Has.Member("SubKey " & i)
    Next i
End Sub
'
Public Sub TestGetSubKeyNamesEmpty()
    Dim s() As String
    s = rk.GetSubKeyNames
    
    Assert.IsFalse cArray.IsNull(s), "Should not be null."
    Assert.AreEqual 0, cArray.GetLength(s), "Should be an empty array."
End Sub

Public Sub TestDeleteSubKeyTree()
    Dim i As Long
    Dim j As Long
    For i = 1 To 5
        RegCreateKeyExA hKey, "SubKey\Child " & i, 0, vbNullString, REG_OPTION_NON_VOLATILE, &HF003F, 0, j, 0
        RegCloseKey j
    Next i
    
    rk.DeleteSubKeyTree "SubKey"
    Assert.IsNothing rk.OpenSubKey("SubKey")
End Sub

Public Sub TestCreateExistingSubKey()
    Dim hk As Long
    RegCreateKeyExA hKey, "SubKey", 0, vbNullString, REG_OPTION_NON_VOLATILE, &HF003F, 0, hk, 0
    RegCloseKey hk
    
    Dim NewKey As RegistryKey
    Set NewKey = rk.CreateSubKey("SubKey")
    Assert.IsNotNothing NewKey
    Assert.AreEqual rk.Name & "\SubKey", NewKey.Name
    NewKey.CloseKey
End Sub

Public Sub TestEnumerateSubKeys()
    Dim i As Long
    Dim j As Long
    For i = 1 To 5
        RegCreateKeyExA hKey, "SubKey " & i, 0, vbNullString, REG_OPTION_NON_VOLATILE, &HF003F, 0, j, 0
        RegCloseKey j
    Next i

    Dim List As New ArrayList
    
    Dim NewKey As RegistryKey
    
End Sub

Public Sub TestEnumerateValueNames()
    Dim List As New ArrayList
    Dim i As Long
    For i = 1 To 10
        RegSetValueExA hKey, "Value " & i, 0, REG_DWORD, VarPtr(i), 4
        List.Add "Value " & i
    Next i

    Dim Name As Variant
    For Each Name In rk.GetEnumerator(detValues)
        Assert.IsTrue List.Contains(Name)
        List.Remove Name
    Next Name
    
    Assert.AreEqual 0, List.Count, "Should be no items remaining."
End Sub

Public Sub TestEnumerateValueEntries()
    Dim List As New Hashtable
    Dim i As Long
    For i = 1 To 10
        RegSetValueExA hKey, "Value " & i, 0, REG_DWORD, VarPtr(i), 4
        List.Add "Value " & i, i
    Next i

    Dim Entry As DictionaryEntry
    For Each Entry In rk.GetEnumerator
        Assert.IsTrue List.Contains(Entry.Key)
        Assert.AreEqual List(Entry.Key), Entry.Value
        List.Remove Entry.Key
    Next Entry

    Assert.AreEqual 0, List.Count, "Should be no items remaining."
End Sub

Public Sub TestEnumerateKeyNames()
    Dim List As New ArrayList
    Dim i As Long
    For i = 1 To 10
        Dim j As Long
        RegCreateKeyExA hKey, "SubKey " & i, 0, vbNullString, REG_OPTION_NON_VOLATILE, &HF003F, 0, j, 0
        RegCloseKey j
        List.Add "SubKey " & i
    Next i

    Dim Name As Variant
    For Each Name In rk.GetEnumerator(detKeys)
        Assert.IsTrue List.Contains(Name)
        List.Remove Name
    Next Name
    
    Assert.AreEqual 0, List.Count, "Should be no items remaining."
End Sub



'Public Sub TestRemote()
'    Dim rk As RegistryKey
'    Dim s() As String
'    Dim i As Long
'
'    Set rk = RegistryKey.OpenRemoteBaseKey(CurrentUser, "Soze")
'    Assert.IsNotNothing rk
'    s = rk.GetSubKeyNames
'    For i = 0 To UBound(s)
'        WriteLine s(i)
'    Next i
'    rk.CloseKey
'End Sub


