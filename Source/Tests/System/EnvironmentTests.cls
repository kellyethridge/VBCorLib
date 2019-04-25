VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnvironmentTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EnvironmentTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Const TestVariable As String = "VBCorLibTests_Set"

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ParseCommandLineArgs_WithValue_ReturnsExpected"
            Test.Use "", Array()
            Test.Use "a\b", Array("a\b")
            Test.Use "a b", Array("a", "b")
            Test.Use "a b c", Array("a", "b", "c")
            Test.Use "a ""b c""", Array("a", "b c")
            Test.Use """a b"" c", Array("a b", "c")
            Test.Use "\""a b\"" c", Array("""a", "b""", "c")
            Test.Use "\\a \\ b", Array("\\a", "\\", "b")
            Test.Use """a""", Array("a")
            Test.Use """a"" ""b"" c", Array("a", "b", "c")
            Test.Use """a""""b""", Array("a""b")
            Test.Use """a""""b\""", Array("a""b""")
            Test.Use "a\\""b", Array("a\\b")
            Test.Use "ab\", Array("ab\")
            Test.Use "ab\\", Array("ab\\")
            Test.Use "\ '", Array("\", "'")
    End Select
End Sub

Public Sub CommandLine_WhenCalled_IsNotEmpty()
    Dim Actual As String
    
    Actual = Environment.CommandLine
    
    Assert.That Actual, Iz.Not.Empty
End Sub

Public Sub MachineName_WhenCalled_IsNotEmpty()
    Dim Actual As String
    
    Actual = Environment.MachineName
    
    Assert.That Actual, Iz.Not.Empty
End Sub

Public Sub OSVersion_WhenCalled_ReturnsVersion()
    Dim Actual As OperatingSystem
    
    Set Actual = Environment.OSVersion
    
    ' I expect tests to be run on Windows 7 or newer, so this
    ' has a pretty good chance of always succeeding.
    Assert.That Actual.Platform, Iz.EqualTo(PlatformID.Win32NT)
End Sub

Public Sub SystemDirectory_WhenCalled_ShouldNotBeEmpty()
    Assert.That Environment.SystemDirectory, Iz.Not.Empty
End Sub

Public Sub ExpandEnvironmentVariables_WhenCalled_ReturnsExpandedVariable()
    Dim Actual As String
    
    Actual = Environment.ExpandEnvironmentVariables("Path=%PATH%")
    
    Assert.That Actual, Iz.EqualTo("Path=" & Environ("PATH"))
End Sub

Public Sub Username_WhenCalled_ShouldNotBeEmpty()
    WriteLine "'" & Environment.UserName & "'"
    Assert.That Environment.UserName, Iz.Not.Empty
End Sub

Public Sub UserDomainName_WhenCalled_ShouldNotBeEmpty()
    WriteLine "UserDomainName: " & Environment.UserDomainName
    Assert.That Environment.UserDomainName, Iz.Not.Empty
End Sub

Public Sub GetEnvironmentVariable_WhenCalled_ReturnsVariable()
    Dim Process As String
    Dim Machine As String
    Dim User    As String
    
    Process = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Process)
    Machine = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine)
    User = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.User)
    
    Assert.That Process, Iz.EqualTo(Environ("PATH")), "#1"
    Assert.That Process, Iz.StringContaining(Machine), "#2"
    Assert.That User, Iz.Not.Empty, "#3"
End Sub

Public Sub GetCommandLineArgs_WhenCalled_ReturnsArgs()
    Dim Actual() As String
    
    Actual = Environment.GetCommandLineArgs
    
    Assert.That CorArray.Length(Actual), Iz.EqualTo(2)
    Assert.That Actual(0), Iz.StringEnding("vb6.exe").IgnoreCase
    Assert.That Actual(1), Iz.StringEnding("corlib.solution.vbg").IgnoreCase
End Sub

Public Sub GetFolderPath_WithWindows_ReturnsWindowsPath()
    Dim Actual As String
    
    Actual = Environment.GetFolderPath(SpecialFolder.Windows)
    
    Assert.That Actual, Iz.EqualTo("c:\windows").IgnoreCase
End Sub

Public Sub GetLogicalDrives_WhenCalled_IncludesCDrive()
    Dim Actual() As String
    
    Actual = Environment.GetLogicalDrives
    
    Assert.That Actual, Has.Member("c:\").IgnoreCase
End Sub

Public Sub GetEnvironmentVariables_WithMachine_ReturnsVariablesForMachine()
    Dim Variables As IDictionary
    
    Set Variables = Environment.GetEnvironmentVariables(Machine)
    
    Dim Entry As DictionaryEntry
    For Each Entry In Variables
        Debug.Print Entry.Key & ": " & Entry.Value
    Next
    
    Assert.That Variables.Keys, Has.Member("path").IgnoreCase
    ' this could fail on other developer machines
    Assert.That Environ("Path"), Iz.StringContaining(Variables("Path"))
End Sub

Public Sub GetEnvironmentVariables_WithProcess_ReturnsVariablesForProcess()
    Dim Variables As IDictionary
    
    Set Variables = Environment.GetEnvironmentVariables
    
    Dim Entry As DictionaryEntry
    For Each Entry In Variables
        Debug.Print Entry.Key & ": " & Entry.Value
    Next
    
    Assert.That Variables.Keys, Has.Member("path").IgnoreCase
End Sub

Public Sub SetEnvironmentVariable_WhenSetForUser_SetsVariableForUserOnly()
    ClearEnvironmentVariable
    
    Environment.SetEnvironmentVariable TestVariable, "ForUser", User
    
    Assert.That Environment.GetEnvironmentVariable(TestVariable, User), Iz.EqualTo("ForUser"), "#1"
    Assert.That Environment.GetEnvironmentVariable(TestVariable, Process), Iz.Not.EqualTo("ForUser"), "#2"
    Assert.That Environment.GetEnvironmentVariable(TestVariable, Machine), Iz.Not.EqualTo("ForUser"), "#3"
End Sub

Public Sub SetEnvironmentVariable_WhenSetForMachine_SetsVariableForMachineAndUser()
    ClearEnvironmentVariable
    
    Environment.SetEnvironmentVariable TestVariable, "ForAll", Machine
    
    Assert.That Environment.GetEnvironmentVariable(TestVariable, Machine), Iz.EqualTo("ForAll"), "#2"
    Assert.That Environment.GetEnvironmentVariable(TestVariable, User), Iz.Not.EqualTo("ForAll"), "#1"
    Assert.That Environment.GetEnvironmentVariable(TestVariable, Process), Iz.Not.EqualTo("ForAll"), "#3"
End Sub

Public Sub ParseCommandLineArgs_WithValue_ReturnsExpected(ByVal LineToParse As String, ByRef Expected As Variant)
    Dim Actual() As String
    
    Actual = Environment.ParseCommandLineArgs(LineToParse)
        
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Private Sub ClearEnvironmentVariable()
    Environment.SetEnvironmentVariable TestVariable, "", User
    Environment.SetEnvironmentVariable TestVariable, "", Machine
    Environment.SetEnvironmentVariable TestVariable, "", Process
End Sub


