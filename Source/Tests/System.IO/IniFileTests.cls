VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniFileTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' IniFileTests
'
Option Explicit
Implements ICategorizable

Private mBuilder As IOHelper

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub Setup()
    Set mBuilder = New IOHelper
End Sub

Public Sub Teardown()
    Set mBuilder = Nothing
End Sub

Public Sub GetString_WhenCalled_ReturnsStoredValue()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    
    Ini.SetValue "Test", "Value", "Hello"
    
    Assert.That Ini.GetString("Test", "Value"), Iz.EqualTo("Hello")
End Sub

Public Sub GetString_WithNonExistingValue_ReturnsEmpty()
    Dim IniFile As IniFile
    Set IniFile = MakeIniFile
    
    Assert.That IniFile.GetString("Test", "Value"), Iz.Empty
End Sub

Public Sub GetString_WhenMissingWithDefaultValue_ReturnsDefaultValue()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    
    Assert.That Ini.GetString("Test", "Value", "Default"), Iz.EqualTo("Default")
End Sub

Public Sub GetString_WithEmptyValue_ReturnsEmptyValue()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    
    Ini.SetValue "Test", "Value", ""
    
    Assert.That Ini.GetString("Test", "Value", "Default"), Iz.Empty
End Sub

Public Sub GetInt32_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", 1
    Ini.SetValue "Test", "Invalid", "abc"
    
    Assert.That Ini.GetInt32("Test", "Value"), Iz.EqualTo(1)
    Assert.That Ini.GetInt32("Test", "Missing", 2), Iz.EqualTo(2)
    Assert.That Ini.GetInt32("Test", "Invalid", 3), Iz.EqualTo(3)
End Sub

Public Sub GetInt16_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", 1
    Ini.SetValue "Test", "Invalid", "abc"
    
    Assert.That Ini.GetInt16("Test", "Value"), Iz.EqualTo(1)
    Assert.That Ini.GetInt16("Test", "Missing", 2), Iz.EqualTo(2)
    Assert.That Ini.GetInt16("Test", "Invalid", 3), Iz.EqualTo(3)
End Sub

Public Sub GetByte_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", 1
    Ini.SetValue "Test", "Invalid", "abc"
    Ini.SetValue "Test", "TooBig", 256
    
    Assert.That Ini.GetByte("Test", "Value"), Iz.EqualTo(1)
    Assert.That Ini.GetByte("Test", "Missing", 2), Iz.EqualTo(2)
    Assert.That Ini.GetByte("Test", "Invalid", 3), Iz.EqualTo(3)
    Assert.That Ini.GetByte("Test", "TooBig", 4), Iz.EqualTo(4)
End Sub

Public Sub GetBoolean_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", True
    Ini.SetValue "Test", "Invalid", "abc"
    Ini.SetValue "Test", "Zero", 0
    Ini.SetValue "Test", "NonZero", 500
    
    Assert.That Ini.GetBoolean("Test", "Value"), Iz.True, "#1"
    Assert.That Ini.GetBoolean("Test", "Missing"), Iz.False, "#2"
    Assert.That Ini.GetBoolean("Test", "Invalid", True), Iz.True, "#3"
    Assert.That Ini.GetBoolean("Test", "Zero", True), Iz.False, "#4"
    Assert.That Ini.GetBoolean("Test", "NonZero"), Iz.True
End Sub

Public Sub GetDate_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", #1/1/2001 8:00:00 AM#
    Ini.SetValue "Test", "Invalid", "abc"
    
    Assert.That Ini.GetDate("Test", "Value"), Iz.EqualTo(#1/1/2001 8:00:00 AM#), "#1"
    Assert.That Ini.GetDate("Test", "Invalid", #1/1/2001#), Iz.EqualTo(#1/1/2001#), "#2"
End Sub

Public Sub GetDateTime_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", Cor.NewDate("2001-01-01 8:00:00AM")
    Ini.SetValue "Test", "Invalid", "abc"
    
    Assert.That Ini.GetDateTime("Test", "Value"), Equals(Cor.NewDate(#1/1/2001 8:00:00 AM#)), "#1"
    Assert.That Ini.GetDateTime("Test", "Invalid", Cor.NewDateTime(2001, 1, 1)), Equals(Cor.NewDateTime(2001, 1, 1)), "#2"
    Assert.That Ini.GetDateTime("Test", "Missing"), Equals(CorDateTime.MinValue), "#3"
End Sub

Public Sub GetTimeSpan_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", TimeSpan.FromHours(5)
    Ini.SetValue "Test", "Invalid", "abc"
    Ini.SetValue "Test", "Date", #1/1/2001 8:00:00 AM#
    
    Assert.That Ini.GetTimeSpan("Test", "Value"), Equals(TimeSpan.FromHours(5)), "#1"
    Assert.That Ini.GetTimeSpan("Test", "Invalid", TimeSpan.FromHours(1)), Equals(TimeSpan.FromHours(1)), "#2"
    Assert.That Ini.GetTimeSpan("Test", "Missing"), Equals(TimeSpan.Zero), "#3"
    Assert.That Ini.GetTimeSpan("Test", "Date"), Equals(TimeSpan.FromHours(8)), "#4"
End Sub

Public Sub GetDouble_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", 1.1
    Ini.SetValue "Test", "Invalid", "abc"
    
    Assert.That Ini.GetDouble("Test", "Value"), Iz.EqualTo(1.1)
    Assert.That Ini.GetDouble("Test", "Missing", 2.2), Iz.EqualTo(2.2)
    Assert.That Ini.GetDouble("Test", "Invalid", 3.3), Iz.EqualTo(3.3)
End Sub

Public Sub GetSingle_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", 1.1
    Ini.SetValue "Test", "Invalid", "abc"
    
    Assert.That Ini.GetSingle("Test", "Value"), Iz.EqualTo(1.1!)
    Assert.That Ini.GetSingle("Test", "Missing", 2.2), Iz.EqualTo(2.2!)
    Assert.That Ini.GetSingle("Test", "Invalid", 3.3), Iz.EqualTo(3.3!)
End Sub

Public Sub GetDecimal_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", 1.1
    Ini.SetValue "Test", "Invalid", "abc"
    
    Assert.That Ini.GetDecimal("Test", "Value"), Iz.EqualTo(1.1)
    Assert.That Ini.GetDecimal("Test", "Missing", 2.2), Iz.EqualTo(2.2)
    Assert.That Ini.GetDecimal("Test", "Invalid", 3.3), Iz.EqualTo(3.3)
End Sub

Public Sub GetCurrency_WhenCalled_ReturnsExpected()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Test", "Value", 1.1
    Ini.SetValue "Test", "Invalid", "abc"
    
    Assert.That Ini.GetCurrency("Test", "Value"), Iz.EqualTo(1.1)
    Assert.That Ini.GetCurrency("Test", "Missing", 2.2), Iz.EqualTo(2.2)
    Assert.That Ini.GetCurrency("Test", "Invalid", 3.3), Iz.EqualTo(3.3)
End Sub

Public Sub GetSectionNames_WhenCalled_ReturnsNames()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Section1", "Value1", 1
    Ini.SetValue "Section2", "Value1", 2
    Ini.SetValue "Section9", "Individual", 11
    
    Dim Names() As String
    Names = Ini.GetSectionNames
    
    Assert.That Names, Iz.EquivalentTo(Array("Section1", "Section2", "Section9"))
End Sub

Public Sub GetSectionNames_WithNoSections_ReturnsEmptyArray()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    
    Dim Names() As String
    Names = Ini.GetSectionNames
    
    Assert.That Names, Iz.Empty
End Sub

Public Sub DeleteKey_WhenCalled_DeletesValueInSection()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Section1", "Value1", 1
    Ini.SetValue "Section1", "Value2", 2
    Ini.SetValue "Section2", "Value1", 3
    
    Ini.DeleteKey "Section1", "Value1"
    
    Assert.That Ini.GetInt32("Section1", "Value1"), Iz.EqualTo(0)
    Assert.That Ini.GetInt32("Section1", "Value2"), Iz.EqualTo(2)
    Assert.That Ini.GetInt32("Section2", "Value1"), Iz.EqualTo(3)
End Sub

Public Sub DeleteSection_WhenCalled_DeletesSection()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Section1", "Value1", 1
    Ini.SetValue "Section2", "Value1", 2
    
    Ini.DeleteSection "Section2"
    
    Assert.That Ini.GetInt32("Section1", "Value1"), Iz.EqualTo(1)
    Assert.That Ini.GetInt32("Section2", "Value1"), Iz.EqualTo(0)
End Sub

Public Sub GetValues_WhenCalled_ReturnsValuesInSection()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Section1", "Value1", 1
    Ini.SetValue "Section1", "Value2", 2
    Ini.SetValue "Section9", "Individual", 11

    Dim Values As IDictionary
    Set Values = Ini.GetValues("Section1")
    
    Assert.That Values.Count, Iz.EqualTo(2)
    Assert.That Values.Contains("Value1"), Iz.True
    Assert.That Values.Contains("Value2"), Iz.True
    Assert.That Values("Value1"), Iz.EqualTo("1")
    Assert.That Values("Value2"), Iz.EqualTo("2")
End Sub

Public Sub GetValues_WithNonExistingSection_ReturnsEmptyValues()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Section1", "Value1", 1
    Ini.SetValue "Section1", "Value2", 2
    Ini.SetValue "Section9", "Individual", 11

    Dim Values As IDictionary
    Set Values = Ini.GetValues("Missing")
    
    Assert.That Values.Count, Iz.EqualTo(0)
End Sub

Public Sub GetKeyNames_WhenCalled_ReturnsNames()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Section1", "Value1", 1
    Ini.SetValue "Section1", "Value2", 2
    Ini.SetValue "Section2", "Value3", 3
    
    Dim Names() As String
    Names = Ini.GetKeyNames("Section1")
    
    Assert.That Names, Iz.EquivalentTo(Array("Value1", "Value2"))
End Sub

Public Sub GetKeyNames_WithNonExistingSection_ReturnsEmptyArray()
    Dim Ini As IniFile
    Set Ini = MakeIniFile
    Ini.SetValue "Section1", "Value1", 1
    Ini.SetValue "Section1", "Value2", 2
    Ini.SetValue "Section2", "Value3", 3

    Dim Names() As String
    Names = Ini.GetKeyNames("Missing")
    
    Assert.That Names, Iz.Empty
End Sub

Private Function MakeIniFile() As IniFile
    Set MakeIniFile = Cor.NewIniFile(mBuilder.CreateFile("File.ini"))
    MakeIniFile.AutoFlush = True
End Function

