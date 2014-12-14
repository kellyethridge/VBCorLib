VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PathTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestPath
'
Option Explicit
Implements ITestCaseSource

Private Declare Function GetLongPathNameA Lib "kernel32.dll" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

Private mOriginalDir As String

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ChangeExtension_WithValue_ReturnsExpected"
            Test.Use "Filename.txt", "rec", "Filename.rec"
            Test.Use "c.:hello", "", "c.:hello."
            Test.Use "hello.txt", "", "hello."
            Test.Use "hello", "d.d", "hello.d.d"
            Test.Use "h.ello", "d.d", "h.d.d"
            Test.Use "hello", ".txt", "hello.txt"
            Test.Use "hello", "txt", "hello.txt"
            Test.Use "", "", ""
            Test.Use "hello.", ".txt", "hello.txt"
            Test.Use "hello\", "txt", "hello\.txt"
            Test.Use "c:\hello", "txt", "c:\hello.txt"
            Test.Use ".txt", "def", ".def"
            
        Case "IsPathRooted_WithValue_ReturnsExpected"
            Test.Use "c:\", True
            Test.Use "\folder", True
            Test.Use "\\share", True
            Test.Use "//share", True
            Test.Use "Abc\file", False
            Test.Use ":\folder", False
            Test.Use "", False
            
        Case "HasExtension_WithValue_ReturnsExpected"
            Test.Use "file.txt", True
            Test.Use ".a", True
            Test.Use ".", False
            Test.Use ".\file", False
            Test.Use "file.", False
            Test.Use "file", False
            
        Case "Combine_WithValues_ReturnsExpected"
            Test.Use "a", "\b", "\b"
            Test.Use "a/", "b", "a/b"
            Test.Use "a", "b", "a\b"
            Test.Use "a:", "b", "a:b"
            Test.Use "c:\", "*.txt", "c:\*.txt"
            Test.Use "", "*.txt", "*.txt"
            Test.Use "", "", ""
            Test.Use "a", "", "a"
            Test.Use "", "b", "b"
            
        Case "GetDirectoryName_WithValue_ReturnsExpected"
            Test.Use "c:\abc", "c:\"
            Test.Use "c:\abc\", "c:\abc"
            Test.Use "c:\", ""
            Test.Use "abc", ""
            Test.Use "abc\", "abc"
            Test.Use "\\server\share", ""
            Test.Use "\\server\share\", "\\server\share"
            Test.Use "\\\\server\\share\\", "\\server\share"
            Test.Use "C:\Documents and Settings\Kelly\Local Settings\Temp\tmpC24.tmp", "C:\Documents and Settings\Kelly\Local Settings\Temp"
            Test.Use "c:\....\abc", "c:\.."
            Test.Use "c:\....\abc\", "c:\..\abc"
            
        Case "GetExtension_WithValue_ReturnsExpected"
            Test.Use "file.txt", ".txt"
            Test.Use "file", ""
            Test.Use ".", ""
            Test.Use "..", ""
            Test.Use "c.:txt", ""
            Test.Use "", ""
            Test.Use "c:\abc.txt\ab", ""
            
        Case "GetFileName_WithValue_ReturnsExpected"
            Test.Use "file.txt", "file.txt"
            Test.Use "c:\file.txt", "file.txt"
            Test.Use "c:\", ""
            Test.Use ":file.txt", "file.txt"
            Test.Use ".txt", ".txt"
            
        Case "GetFileNameWithoutExtension_WithValue_ReturnsExpected"
            Test.Use "file.txt", "file"
            Test.Use "c:\file.txt", "file"
            Test.Use "c:\", ""
            Test.Use "c:file.txt", "file"
            Test.Use "file", "file"
            Test.Use "c:file", "file"
            Test.Use "c:.file", ""
            Test.Use "", ""
            Test.Use "c:\file.txt.txt", "file.txt"
            Test.Use "file.txt.txt", "file.txt"
            Test.Use "file.txt\", ""
            Test.Use ".txt", ""

        Case "GetFullPath_WithValue_ReturnsExpected"
            Test.Use "file.txt", "c:\windows\file.txt"
            Test.Use "\file.txt", "c:\file.txt"
            Test.Use "\\server\share", "\\server\share"
            Test.Use "\\server\share\", "\\server\share\"
            Test.Use "c:\abc.txt", "c:\abc.txt"
            Test.Use "c\abc.txt", "c:\windows\c\abc.txt"
            Test.Use "a\b...\a", "c:\windows\a\b\a"
            Test.Use "a\...b\a", "c:\windows\a\...b\a"
            Test.Use "c:\", "c:\"
            Test.Use "c:\\\", "c:\"
            Test.Use "\\\\server\\\share", "\\server\share"
            Test.Use "\\\\server\\\share\\\", "\\server\share\"
            Test.Use "c:\abc\def\ghi\..\..\jkl", "c:\abc\jkl"
            Test.Use "..\jkl", "c:\jkl"
            
    End Select
End Sub

Public Sub Setup()
    mOriginalDir = CurDir
    ChDir "c:\windows"
End Sub

Public Sub Teardown()
    ChDir mOriginalDir
End Sub

Public Sub ChangeExtension_WithValue_ReturnsExpected(ByVal PathToChange As String, ByVal NewExtension As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Path.ChangeExtension(PathToChange, NewExtension)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ChangeExtension_WithNullString_RemovesExtension()
    Dim Actual As String
    
    Actual = Path.ChangeExtension("hello.txt", vbNullString)
    
    Assert.That Actual, Iz.EqualTo("hello")
End Sub

Public Sub IsPathRooted_WithValue_ReturnsExpected(ByVal PathToCheck As String, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Path.IsPathRooted(PathToCheck)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub HasExtension_WithValue_ReturnsExpected(ByVal PathToCheck As String, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Path.HasExtension(PathToCheck)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Combine_WithValues_ReturnsExpected(ByVal Path1 As String, ByVal Path2 As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Path.Combine(Path1, Path2)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDirectoryName_WithValue_ReturnsExpected(ByVal PathToUse As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Path.GetDirectoryName(PathToUse)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetExtension_WithValue_ReturnsExpected(ByVal PathToUse As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Path.GetExtension(PathToUse)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetFileName_WithValue_ReturnsExpected(ByVal PathToUse As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Path.GetFileName(PathToUse)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetFileNameWithoutExtension_WithValue_ReturnsExpected(ByVal PathToUse As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Path.GetFileNameWithoutExtension(PathToUse)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetFullPath_WithValue_ReturnsExpected(ByVal PathToUse As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Path.GetFullPath(PathToUse)
    
    Assert.That Actual, Iz.EqualTo(Expected).IgnoreCase
End Sub

Public Sub TestGetFullPathFixSlashes()
    Assert.AreEqual "\\server\share\file.txt", Path.GetFullPath("//server/share/file.txt")
End Sub

Public Sub TestGetFullPathInvalidUNCShareWithoutSlash()
    On Error GoTo errTrap
    Call Path.GetFullPath("\\share")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub TestGetFullPathInvalidUNCShareWithSlash()
    On Error GoTo errTrap
    Call Path.GetFullPath("\\share\")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub TestGetFullPathOfEmptyPath()
    On Error GoTo errTrap
    Call Path.GetFullPath("")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub TestGetFullPathOfSpaces()
    On Error GoTo errTrap
    Call Path.GetFullPath("  ")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub TestGetFullPathMultipleColons()
    On Error GoTo errTrap
    Path.GetFullPath "c:\a:"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is NotSupportedException, "Wrong exception type."
End Sub

Public Sub TestGetPathRoot()
    Assert.AreEqual "\", Path.GetPathRoot("\mydir\")
    Assert.AreEqual "c:", Path.GetPathRoot("c:")
    Assert.AreEqual "", Path.GetPathRoot("file.txt")
    Assert.AreEqual "\\server", Path.GetPathRoot("\\server")
    Assert.AreEqual "\\server\", Path.GetPathRoot("\\server\")
    Assert.AreEqual "\\server\share", Path.GetPathRoot("\\server\share")
    Assert.AreEqual "\\server\share", Path.GetPathRoot("\\server\share\dir")
    Assert.AreEqual "\\", Path.GetPathRoot("\\")
    Assert.AreEqual "\\", Path.GetPathRoot("\\\")
    Assert.AreEqual "\\.\", Path.GetPathRoot("\\.\")
    Assert.AreEqual "\\a", Path.GetPathRoot("\\\a")
End Sub

Public Sub TestGetTempFileName()
    Dim s As String
    s = Path.GetTempFileName
    Assert.That "", Iz.Not.EqualTo(s)
    Assert.AreEqual 0, FileLen(s)
    Kill s
    
    Assert.AreEqual GetTempFolder, Path.GetDirectoryName(s)
End Sub

    ' the output is user specific, so visual inspection is suggested.
Public Sub TestGetTempPath()
    Assert.AreEqual GetTempFolder & "\", Path.GetTempPath
End Sub



Private Function GetTempFolder() As String
    Dim Tmp As String
    Tmp = Environ$("Tmp")
    If Len(Tmp) = 0 Then Tmp = Environ$("Temp")
    If Len(Tmp) = 0 Then Tmp = Environ$("UserProfile")
    If Len(Tmp) = 0 Then Tmp = Environ$("WinDir")
    Tmp = GetLongPathName(Tmp)

    GetTempFolder = Tmp
End Function

Private Function GetLongPathName(ByVal s As String) As String
    Dim ret As String
    ret = String$(1024, 0)
    Dim i As Long
    i = GetLongPathNameA(s, ret, Len(ret))
    GetLongPathName = Left$(ret, i)
End Function


