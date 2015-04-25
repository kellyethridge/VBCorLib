VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestEnvironment
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

' These don't really test the Environment class since the results
' will vary from machine to machine. Instead they simply display
' the results for the current machine for visual inspection.
Public Sub TestFixtureSetup()
    WriteLine ""
    WriteLine "***** ENVIRONMENT CLASS *****"
End Sub

Public Sub TestCommandLine()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "Command Line: " & Environment.CommandLine
End Sub

Public Sub TestMachineName()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "Machine Name: " & Environment.MachineName
End Sub

Public Sub TestOSVersion()
    Dim os As OperatingSystem
    
    Set os = Environment.OSVersion
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "OS: " & os.ToString
    WriteLine "Version: " & os.Version.ToString
    WriteLine "Major Revision: " & os.Version.MajorRevision
    WriteLine "Minor Revision: " & os.Version.MinorRevision
End Sub

Public Sub TestSystemDirectory()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "System Directory: " & Environment.SystemDirectory
End Sub

Public Sub TestTickCount()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "TickCount: " & Environment.TickCount
End Sub

Public Sub TestUserName()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "UserName: " & Environment.UserName
End Sub

Public Sub TestUserDomainName()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "UserDomainName: " & Environment.UserDomainName
End Sub

Public Sub TestUserInteractive()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "UserInteractive: " & Environment.UserInteractive
End Sub

Public Sub TestWorkingSet()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "Returns 0 for 95, 98 and ME installations."
    WriteLine "WorkingSet (Memory Usage):" & Environment.WorkingSet
End Sub

Public Sub TestExpandEnvironmentVariables()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "--- ExpandingEnvironmentVariables ---"
    WriteLine "'My OS is %OS%' becomes: " & Environment.ExpandEnvironmentVariables("My OS is %OS%")
End Sub

Public Sub TestGetCommandLineArgs()
    Dim s() As String
    Dim i As Long
    
    s = Environment.GetCommandLineArgs
    
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "--- GetCommandLingArgs ---"
    WriteLine "Should have no empty args unless quoted."
    WriteLine "Count: " & cArray.GetLength(s)
    WriteLine "Program: " & s(0)
    For i = 1 To UBound(s)
        WriteLine "Arg(" & i & ") = " & s(i)
    Next i
End Sub

Public Sub TestGetEnvironmentVariable()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "GetEnvironmentVariable (os): " & Environment.GetEnvironmentVariable("os")
End Sub

Public Sub TestGetEnvironmentVariables()
    Dim ht As Hashtable
    Dim i As Long
    
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "--- TestGetEnvironmentVariables ---"
    
    Set ht = Environment.GetEnvironmentVariables
    WriteLine "Count: " & ht.Count
    
    Dim de As DictionaryEntry
    For Each de In ht
        i = i + 1
        WriteLine "Variable(" & i & "): " & de.Key & " = " & de.Value
    Next de
End Sub

Public Sub TestGetFolderPath()
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "--- TestGetFolderPath ---"
    WriteLine "My Documents: " & Environment.GetFolderPath(Personal)
    WriteLine "Desktop: " & Environment.GetFolderPath(DesktopDirectory)
    WriteLine "Windows: " & Environment.GetFolderPath(Windows)
    WriteLine "System: " & Environment.GetFolderPath(SystemDir)
End Sub

Public Sub TestGetLogicalDrives()
    Dim s() As String
    Dim i As Long
    
    s = Environment.GetLogicalDrives
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "--- GetLogicalDrives ---"
    For i = 0 To UBound(s)
        WriteLine "Drive - " & s(i)
    Next i
End Sub

Public Sub TestUserGetEnvironmentVariables()
    Dim ht As Hashtable
    Dim i As Long
    
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "--- TestUserGetEnvironmentVariables ---"
    
    Set ht = Environment.GetEnvironmentVariables(User)
    WriteLine "Count: " & ht.Count
    
    Dim de As DictionaryEntry
    For Each de In ht
        i = i + 1
        WriteLine "Variable(" & i & "): " & de.Key & " = " & de.Value
    Next de
End Sub

Public Sub TestMachineGetEnvironmentVariables()
    Dim ht As Hashtable
    Dim i As Long
    
    WriteLine ""
    WriteLine "--- VISUAL INSPECTION ---"
    WriteLine "--- TestMachineGetEnvironmentVariables ---"
    
    Set ht = Environment.GetEnvironmentVariables(Machine)
    WriteLine "Count: " & ht.Count
    
    Dim de As DictionaryEntry
    For Each de In ht
        i = i + 1
        WriteLine "Variable(" & i & "): " & de.Key & " = " & de.Value
    Next de
End Sub

Public Sub TestSetEnvironmentVariable()
    Dim v As String

    Environment.SetEnvironmentVariable "VBCorLibTests_Set", "Succeed"

    v = Environment.GetEnvironmentVariable("VBCorLibTests_Set")
    Environment.SetEnvironmentVariable "VBCorLibTests_Set", ""
    Assert.AreEqual "Succeed", v
    Assert.AreEqual "", Environment.GetEnvironmentVariable("VBCorLibTests_Set")
End Sub

Public Sub TestParseEmptyCommandArgs()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs("")
    Assert.IsFalse cArray.IsNull(Args), "Should not be null."
    Assert.AreEqual -1, UBound(Args)
End Sub

Public Sub TestParseOneArg()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs("a1")
    Assert.AreEqual 0, UBound(Args)
    Assert.AreEqual "a1", Args(0)
End Sub

Public Sub TestQuotedArgReturnsNonQuotedArg()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs(Q("a1"))
    Assert.AreEqual 0, UBound(Args)
    Assert.AreEqual "a1", Args(0)
End Sub

Public Sub TestParseTwoArgs()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs("a1 a2")
    Assert.AreEqual 1, UBound(Args)
    Assert.AreEqual "a1", Args(0)
    Assert.AreEqual "a2", Args(1)
End Sub

Public Sub TestTwoQuotedArgsReturnTwoNonQuotedArgs()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs(Q("a1") & " " & Q("a2"))
    Assert.AreEqual 1, UBound(Args)
    Assert.AreEqual "a1", Args(0)
    Assert.AreEqual "a2", Args(1)
End Sub

Public Sub TestTwoQuotedArgsNoSpacesReturnOneNonQuotedArg()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs(Q("a1") & Q("a2"))
    Assert.AreEqual 0, UBound(Args)
    Assert.AreEqual "a1" & Chr$(34) & "a2", Args(0)
End Sub

Public Sub TestEscapedQuote()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs("\" & Chr$(34))
    Assert.AreEqual 0, UBound(Args)
    Assert.AreEqual Chr$(34), Args(0)
End Sub

Public Sub TestTwoQuotedArgsWithLastQuoteEscaped()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs(Q("a1") & Q("a2\"))
    Assert.AreEqual 0, UBound(Args)
    Assert.AreEqual "a1" & Chr$(34) & "a2" & Chr$(34), Args(0)
End Sub

Public Sub TestTwoQuotedArgsWithBackSlashSeparator()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs(Q("a1") & "\" & Q("a2"))
    Assert.AreEqual 0, UBound(Args)
    Assert.AreEqual "a1" & Chr$(34) & "a2", Args(0)
End Sub

Public Sub TestQuotedArguments()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs(Q("a 1"))
    Assert.AreEqual "a 1", Args(0)
End Sub

Public Sub TestEscapedBackSlashBeforeQuote()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs("File\\" & Chr$(34) & "abc")
    Assert.AreEqual "File\\abc", Args(0)
End Sub

Public Sub TestEndingWithBackSlashes()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs("abc\\")
    Assert.AreEqual "abc\\", Args(0)
End Sub

Public Sub TestBackslashIncluded()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs("abc\def")
    Assert.AreEqual 0, UBound(Args)
    Assert.AreEqual "abc\def", Args(0)
End Sub

Public Sub TestBackslashIsOneArgument()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs("\ '")
    Assert.AreEqual 1, UBound(Args)
    Assert.AreEqual "\", Args(0)
    Assert.AreEqual "'", Args(1)
End Sub

Public Sub TestBackslashAtEnd()
    Dim Args() As String
    Args = Environment.ParseCommandLineArgs("abc\")
    Assert.AreEqual 0, UBound(Args)
    Assert.AreEqual "abc\", Args(0)
End Sub



Private Function Q(ByVal s As String) As String
    Q = Chr$(34) & s & Chr$(34)
End Function
