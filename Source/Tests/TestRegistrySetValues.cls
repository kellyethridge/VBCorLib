VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRegistrySetValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSetRegistryValues
'
Option Explicit
'Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExStr Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Private Key As RegistryKey
Private hKey As Long

Public Sub Setup()
    Set Key = Registry.CurrentUser.CreateSubKey("software\vb and vba program settings\VBCorLibTests\Tests")
    RegOpenKeyExA &H80000001, "software\vb and vba program settings\VBCorLibTests\Tests", 0, 983103, hKey
End Sub

Public Sub TearDown()
    RegCloseKey hKey
    Key.CloseKey
    Registry.CurrentUser.DeleteSubKeyTree "software\vb and vba program settings\VBCorLibTests"
End Sub

Public Sub TestEmptyString()
    Key.SetValue "Empty", ""
    AssertKindAndSize "Empty", StringKind, 1
    
    Dim s As String
    s = " "
    RegQueryValueExStr hKey, "Empty", 0, 0, s, 1
    Assert.AreEqual vbNullChar, s
End Sub

Public Sub TestString()
    Key.SetValue "Value", "Hello"
    
    AssertKindAndSize "Value", StringKind, 6
    
    Dim s As String
    s = String$(6, 0)
    RegQueryValueExStr hKey, "Value", 0, 0, s, 6
    Assert.AreEqual "Hello" & vbNullChar, s
End Sub

Public Sub TestEmptyMultiString()
    Key.SetValue "Empty", cor.NewStrings()
    
    AssertKindAndSize "Empty", MultiStringKind, 1
    
    Dim s As String
    s = " "
    RegQueryValueExStr hKey, "Empty", 0, 0, s, 1
    Assert.AreEqual vbNullChar, s
End Sub

Public Sub TestDWord()
    Key.SetValue "Long", &H7FFFFFFF
    Key.SetValue "Integer", &H7FFF
    Key.SetValue "Byte", CByte(&HFF)
    
    AssertKindAndSize "Long", DWordKind, 4
    AssertKindAndSize "Integer", DWordKind, 4
    AssertKindAndSize "Byte", DWordKind, 4
    
    Dim value As Long
    RegQueryValueExA hKey, "Long", 0, 0, value, 4
    Assert.AreEqual &H7FFFFFFF, value
    
    RegQueryValueExA hKey, "Integer", 0, 0, value, 4
    Assert.AreEqual &H7FFF, value
    
    RegQueryValueExA hKey, "Byte", 0, 0, value, 4
    Assert.AreEqual &HFF, value
End Sub

Public Sub TestQWord()
    Key.SetValue "Value", 12345.6789@
    
    AssertKindAndSize "Value", QWordKind, 8
    
    Dim c As Currency
    RegQueryValueExA hKey, "Value", 0, 0, c, 8
    Assert.AreEqual 12345.6789@, c
End Sub

Public Sub TestBinary()
    Dim Bytes() As Byte
    Bytes = cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0)
    Key.SetValue "Value", Bytes
    
    AssertKindAndSize "Value", BinaryKind, 20
    
    ReDim b(19) As Byte
    RegQueryValueExA hKey, "Value", 0, 0&, b(0), 20&
    
    Assert.AreEqual Bytes, b
End Sub

Public Sub TestStringAsDWord()
    Key.SetValue "Value", "1234", DWordKind

    AssertKindAndSize "Value", DWordKind, 4
    
    Dim value As Long
    RegQueryValueExA hKey, "Value", 0, 0&, value, 4&
    
    Assert.AreEqual 1234, value
End Sub

Public Sub TestDWordAsString()
    Key.SetValue "Value", 12345, StringKind

    AssertKindAndSize "Value", StringKind, 6
    
    Dim value As String
    value = String$(6, 0)
    RegQueryValueExStr hKey, "Value", 0, 0, value, 6
    Assert.AreEqual "12345" & vbNullChar, value
End Sub

Public Sub TestDoubleAsString()
    Key.SetValue "Value", 1234.567
    
    AssertKindAndSize "Value", StringKind, 9
    
    Dim value As String
    value = String$(9, 0)
    RegQueryValueExStr hKey, "Value", 0, 0, value, 9
    Assert.AreEqual "1234.567" & vbNullChar, value
End Sub

Public Sub TestDWordAsQWord()
    Key.SetValue "Value", 12345, QWordKind
    
    AssertKindAndSize "Value", QWordKind, 8

    Dim c As Currency
    RegQueryValueExA hKey, "Value", 0, 0, c, 8
    
    Assert.AreEqual 1.2345, c
End Sub

Public Sub TestQWordAsQWord()
    Key.SetValue "Value", 12345@, QWordKind
    
    AssertKindAndSize "Value", QWordKind, 8
    
    Dim c As Currency
    RegQueryValueExA hKey, "Value", 0, 0, c, 8
    
    Assert.AreEqual 12345@, c
End Sub

Public Sub TestNullMultiSZ()
    Dim MultiSZ() As String
    Key.SetValue "Empty", MultiSZ
    
    AssertKindAndSize "Empty", MultiStringKind, 1
    
    Dim s As String
    s = " "
    RegQueryValueExStr hKey, "Empty", 0, 0, s, 1
    Assert.AreEqual vbNullChar, s
End Sub

Public Sub TestNullBinary()
    Dim Bytes() As Byte
    Key.SetValue "Value", Bytes
    
    AssertKindAndSize "Value", BinaryKind, 0
End Sub

Public Sub TestEmptyBinary()
    Dim Bytes() As Byte
    Bytes = cor.NewBytes
    
    Key.SetValue "Value", Bytes
    
    AssertKindAndSize "Value", BinaryKind, 0
End Sub

Public Sub TestMultiSZ()
    Key.SetValue "Value", cor.NewStrings("one", "two", "three")
    AssertKindAndSize "Value", MultiStringKind, 15
    
    Dim s As String
    s = String$(15, 0)
    RegQueryValueExStr hKey, "Value", 0, 0, s, 15
    Assert.AreEqual "one" & vbNullChar & "two" & vbNullChar & "three" & vbNullChar & vbNullChar, s
End Sub

Public Sub TestExpandSZ()
    Key.SetValue "Value", "Expanded %Path%", ExpandStringKind
    AssertKindAndSize "Value", ExpandStringKind, 16
    
    Dim s As String
    s = String$(15, 0)
    RegQueryValueExStr hKey, "Value", 0, 0, s, 16
    Assert.That s, Iz.EqualTo("Expanded %Path%")
End Sub

Public Sub TestDefaultAsEmptySZ()
    Key.SetValue "", ""
    
    AssertKindAndSize "", StringKind, 1
    
    Dim s As String
    s = " "
    RegQueryValueExStr hKey, "", 0, 0, s, 1
    Assert.AreEqual vbNullChar, s
End Sub

Public Sub TestDefaultAsSZ()
    Key.SetValue "", "My Default"
    
    AssertKindAndSize "", StringKind, Len("My Default") + 1
    
    Dim s As String
    s = String$(Len("My Default") + 1, 0)
    RegQueryValueExStr hKey, "", 0, 0, s, Len(s)
    Assert.AreEqual "My Default" & vbNullChar, s
End Sub

Public Sub TestNameTooLong()
    Dim Name As String
    Name = String$(Registry.MaxValueNameLength + 1, "f")
    
    On Error GoTo errTrap
    Key.SetValue Name, 1
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestValueKind()
    Key.SetValue "DWORD", 1
    Key.SetValue "SZ", "hi"
    Key.SetValue "QWORD", 1@
    Key.SetValue "BINARY", cor.NewBytes(0)
    Key.SetValue "EXPAND", "hi", ExpandStringKind
    Key.SetValue "MULTI", Array("hi", "bye")
    
    Assert.AreEqual RegistryValueKind.DWordKind, Key.GetValueKind("DWORD")
    Assert.AreEqual RegistryValueKind.StringKind, Key.GetValueKind("SZ")
    Assert.AreEqual RegistryValueKind.QWordKind, Key.GetValueKind("QWORD")
    Assert.AreEqual RegistryValueKind.BinaryKind, Key.GetValueKind("BINARY")
    Assert.AreEqual RegistryValueKind.ExpandStringKind, Key.GetValueKind("EXPAND")
    Assert.AreEqual RegistryValueKind.MultiStringKind, Key.GetValueKind("MULTI")
End Sub



Private Sub AssertKindAndSize(ByVal Name As String, ByVal ExpectedKind As RegistryValueKind, ByVal ExpectedSize As Long)
    Dim Kind As RegistryValueKind
    Dim Size As Long
    
    RegQueryValueExA hKey, Name, 0, Kind, ByVal 0&, Size
    Assert.AreEqual ExpectedKind, Kind
    Assert.AreEqual ExpectedSize, Size
End Sub
