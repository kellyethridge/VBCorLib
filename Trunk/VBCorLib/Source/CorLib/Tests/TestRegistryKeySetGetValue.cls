VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRegistryKeySetGetValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRegistryKeySetGetValue
'
Option Explicit

Private Key As RegistryKey

Public Sub Setup()
    Set Key = Registry.CurrentUser.CreateSubKey("software\vb and vba program settings\VBCorLibTests\Tests")
End Sub

Public Sub TearDown()
    Key.CloseKey
    Registry.CurrentUser.DeleteSubKeyTree "software\vb and vba program settings\VBCorLibTests"
End Sub


Public Sub TestMissingNoDefault()
    Assert.IsVBEmpty Key.GetValue("Missing")
End Sub

Public Sub TestString()
    Key.SetValue "Name", "Kelly"
    Assert.AreEqual "Kelly", Key.GetValue("Name")
End Sub

Public Sub TestStringMissingWithDefault()
    Assert.AreEqual "Default", Key.GetValue("Missing", "Default")
End Sub

Public Sub TestLong()
    Key.SetValue "Age", 36
    Assert.AreEqual 36, Key.GetValue("Age")
End Sub

Public Sub TestLongMissingWithDefault()
    Assert.AreEqual 25, Key.GetValue("Missing", 25)
End Sub

Public Sub TestBinary()
    Dim b() As Byte
    
    Key.SetValue "Bytes", cArray.NewArray(ciByte, 1, 2, 3)
    b = Key.GetValue("Bytes")
    Assert.AreEqual 0, LBound(b), "Should have an lbound of 0."
    Assert.AreEqual 2, UBound(b), "Should have a ubound of 2."
    Assert.AreEqual 1, b(0)
    Assert.AreEqual 2, b(1)
    Assert.AreEqual 3, b(2)
End Sub

Public Sub TestBinaryMissingWithDefault()
    Dim b() As Byte
    
    b = Key.GetValue("Missing", cArray.NewArray(ciByte, 3, 2, 1))
    Assert.AreEqual 0, LBound(b), "should have an lbound of 0."
    Assert.AreEqual 2, UBound(b), "should have a ubound of 2."
    Assert.AreEqual 3, b(0)
    Assert.AreEqual 2, b(1)
    Assert.AreEqual 1, b(2)
End Sub

Public Sub TestMultiString()
    Dim s() As String
    
    Key.SetValue "Strings", Split("one,two,three", ",")
    s = Key.GetValue("Strings")
    Assert.AreEqual 0, LBound(s), "should have an lbound of 0."
    Assert.AreEqual 2, UBound(s), "should have a ubound of 2."
    Assert.AreEqual "one", s(0)
    Assert.AreEqual "two", s(1)
    Assert.AreEqual "three", s(2)
End Sub

Public Sub TestMultiStringMissingWithDefault()
    Dim s() As String
    
    s = Key.GetValue("Missing", Split("three,two,one", ","))
    Assert.AreEqual 0, LBound(s), "should have an lbound of 0."
    Assert.AreEqual 2, UBound(s), "should have a ubound of 2."
    Assert.AreEqual "three", s(0)
    Assert.AreEqual "two", s(1)
    Assert.AreEqual "one", s(2)
End Sub

Public Sub TestDefaultDataType()
    Key.SetValue "Double", 1.2
    Assert.AreEqual "1.2", GetSetting("VBCorLibTests", "Tests", "Double")
End Sub

Public Sub TestObjectToString()
    Dim ts As TimeSpan
    
    Set ts = TimeSpan.FromHours(25)
    
    Key.SetValue "Object", ts
    Assert.AreEqual ts.ToString, Key.GetValue("Object")
End Sub

Public Sub TestKeyDefaultValue()
    Key.SetValue "", "My Default Value"
    Assert.AreEqual "My Default Value", Key.GetValue("")
End Sub

Public Sub TestExpandedString()
    Key.SetValue "ExpandedString", "My path is %PATH%", ExpandStringKind
    Assert.AreEqual "My path is %PATH%", Key.GetValue("ExpandedString", , DoNotExpandEnvironmentVariables)
End Sub

Public Sub TestQWord()
    Key.SetValue "QWordAdjustment", 1, QWordKind
    Assert.AreEqual 0.0001@, Key.GetValue("QWordAdjustment")
    
    Key.SetValue "QWordNoAdjustment", 1@, QWordKind
    Assert.AreEqual 1@, Key.GetValue("QWordNoAdjustment")
End Sub

Public Sub TestEmptyBinary()
    Dim b() As Byte
    
    Key.SetValue "EmptyBytes", b
    
    ReDim b(10)
    b = Key.GetValue("EmptyBytes")
    Assert.IsFalse cArray.IsNull(b), "Array should not be null."
    Assert.AreEqual 0, cArray.GetLength(b)
End Sub

Public Sub TestEmptyMultiString()
    Dim s() As String
    
    Key.SetValue "EmptyStrings", s
    
    ReDim s(10)
    s = Key.GetValue("EmptyStrings")
    Assert.IsFalse cArray.IsNull(s), "Array should not be null."
    Assert.AreEqual 0, cArray.GetLength(s)
End Sub

