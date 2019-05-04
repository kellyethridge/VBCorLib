VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniResourceWriterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' IniResourceWriterTests
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

Public Sub NewIniResourceWriter_WhenCalled_ReturnsInstance()
    Dim Writer As IniResourceWriter
    
    Set Writer = Cor.NewIniResourceWriter("abc.ini", "sec1")
    
    Assert.That Writer.FileName, Iz.EqualTo("abc.ini")
    Assert.That Writer.Section, Iz.EqualTo("sec1")
End Sub

Public Sub Generate_WithNoResourcesAdded_AddsEmptySection()
    Dim Writer As IniResourceWriter
    Set Writer = MakeWriter
    
    Writer.Generate
    
    Dim IniFile As IniFile
    Set IniFile = MakeIniFile
    Assert.That IniFile.GetValues("Section1"), Iz.Empty
End Sub

Public Sub Generate_WithResourcesAdded_AddsToSection()
    Dim Writer As IniResourceWriter
    Set Writer = MakeWriter
    
    Writer.AddResource "Value1", 1
    Writer.Generate
    
    Dim IniFile As IniFile
    Set IniFile = MakeIniFile
    Assert.That IniFile.GetInt32("Section1", "Value1"), Iz.EqualTo(1)
End Sub

Public Sub CloseWriter_WhenNotGenerated_GeneratesResource()
    Dim Writer As IniResourceWriter
    Set Writer = MakeWriter
    
    Writer.AddResource "Value1", 1
    Writer.CloseWriter
    
    Dim IniFile As IniFile
    Set IniFile = MakeIniFile
    Assert.That IniFile.GetInt32("Section1", "Value1"), Iz.EqualTo(1)
End Sub

Public Sub AddResource_OnceGenerated_ThrowsInvalidOperationException()
    Dim Writer As IniResourceWriter
    Set Writer = MakeWriter
    Writer.Generate
    
    On Error Resume Next
    Writer.AddResource "Value1", 1
    
    AssertInvalidOperationException Err
End Sub


Private Function MakeWriter() As IniResourceWriter
    Set MakeWriter = Cor.NewIniResourceWriter(mBuilder.CreateFile("File.ini"), "Section1")
End Function

Private Function MakeIniFile() As IniFile
    Set MakeIniFile = Cor.NewIniFile(mBuilder.CreateFile("File.ini"))
    MakeIniFile.AutoFlush = True
End Function
