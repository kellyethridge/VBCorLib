VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestPath
'
Option Explicit
Private Declare Function GetLongPathNameA Lib "kernel32.dll" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long


Public Sub TestChangeExtension()
    Assert.AreEqual "Filename.rec", Path.ChangeExtension("Filename.txt", "rec")
    Assert.AreEqual "c.:hello.", Path.ChangeExtension("c.:hello", "")
    Assert.AreEqual "hello", Path.ChangeExtension("hello.txt", vbNullString)
    Assert.AreEqual "hello.d.d", Path.ChangeExtension("hello", "d.d")
    Assert.AreEqual "h.d.d", Path.ChangeExtension("h.ello", "d.d")
    Assert.AreEqual "h", Path.ChangeExtension("h.ello", vbNullString)
    Assert.AreEqual "hello.txt", Path.ChangeExtension("hello", ".txt")
    Assert.AreEqual "hello.txt", Path.ChangeExtension("hello", "txt")
    Assert.AreEqual "", Path.ChangeExtension("", "")
    Assert.AreEqual "hello.txt", Path.ChangeExtension("hello.", ".txt")
    Assert.AreEqual "hello\.txt", Path.ChangeExtension("hello\", "txt")
    Assert.AreEqual "c:\hello.txt", Path.ChangeExtension("c:\hello", "txt")
    Assert.AreEqual ".def", Path.ChangeExtension(".txt", "def")
End Sub

Public Sub TestRootedPaths()
    Assert.IsTrue Path.IsPathRooted("c:\"), "'c:\' should be rooted."
    Assert.IsTrue Path.IsPathRooted("\folder"), "'\folder' should be rooted."
    Assert.IsTrue Path.IsPathRooted("\\share"), "'\\share' should be rooted."
    Assert.IsTrue Path.IsPathRooted("//share"), "'//share' should be rooted."
End Sub

Public Sub TestNotRootedPaths()
    Assert.IsFalse Path.IsPathRooted("Abc\file"), "'Abc\file' should not be rooted."
    Assert.IsFalse Path.IsPathRooted(":\folder"), "':\folder' should not be rooted."
    Assert.IsFalse Path.IsPathRooted(""), "empty path should not be rooted."
End Sub

Public Sub TestHasExtension()
    Assert.IsTrue Path.HasExtension("file.txt"), "'file.txt' should have an extension."
    Assert.IsTrue Path.HasExtension(".a"), "'.a' should have an extension."
    Assert.IsFalse Path.HasExtension("."), "'.' should not have an extension."
    Assert.IsFalse Path.HasExtension(".\file"), "'.\file' should not have an extension."
    Assert.IsFalse Path.HasExtension("file."), "'file.' should not have an extension."
    Assert.IsFalse Path.HasExtension("file"), "'file' should not have an extension."
End Sub

Public Sub TestCombine()
    Assert.AreEqual "a\b", Path.Combine("a", "b")
    Assert.AreEqual "\b", Path.Combine("a", "\b")
    Assert.AreEqual "a/b", Path.Combine("a/", "b")
    Assert.AreEqual "a:b", Path.Combine("a:", "b")
    Assert.AreEqual "c:\*.txt", Path.Combine("c:\", "*.txt")
    Assert.AreEqual "*.txt", Path.Combine("", "*.txt")
    Assert.AreEqual "", Path.Combine("", "")
    Assert.AreEqual "a", Path.Combine("a", "")
    Assert.AreEqual "b", Path.Combine("", "b")
End Sub

Public Sub TestGetDirectoryNameWithLegalPaths()
    Assert.AreEqual "c:\", Path.GetDirectoryName("c:\abc")
    Assert.AreEqual "c:\abc", Path.GetDirectoryName("c:\abc\")
    Assert.AreEqual "", Path.GetDirectoryName("c:\")
    Assert.AreEqual "", Path.GetDirectoryName("abc")
    Assert.AreEqual "abc", Path.GetDirectoryName("abc\")
    Assert.AreEqual "", Path.GetDirectoryName("abc")
    Assert.AreEqual "", Path.GetDirectoryName("\\server\share")
    Assert.AreEqual "\\server\share", Path.GetDirectoryName("\\server\share\")
    Assert.AreEqual "\\server\share", Path.GetDirectoryName("\\\\server\\share\\")
    Assert.AreEqual "C:\Documents and Settings\Kelly\Local Settings\Temp", Path.GetDirectoryName("C:\Documents and Settings\Kelly\Local Settings\Temp\tmpC24.tmp")
    Assert.AreEqual "c:\..", Path.GetDirectoryName("c:\....\abc")
    Assert.AreEqual "c:\..\abc", Path.GetDirectoryName("c:\....\abc\")
End Sub

Public Sub TestGetExtension()
    Assert.AreEqual ".txt", Path.GetExtension("file.txt")
    Assert.AreEqual "", Path.GetExtension("file")
    Assert.AreEqual ".file", Path.GetExtension(".file")
    Assert.AreEqual "", Path.GetExtension(".")
    Assert.AreEqual "", Path.GetExtension("..")
    Assert.AreEqual "", Path.GetExtension("c.:txt")
    Assert.AreEqual "", Path.GetExtension("")
    Assert.AreEqual "", Path.GetExtension("c:\abc.txt\ab")
End Sub

Public Sub TestGetFileName()
    Assert.AreEqual "file.txt", Path.GetFileName("file.txt")
    Assert.AreEqual "file.txt", Path.GetFileName("c:\file.txt")
    Assert.AreEqual "", Path.GetFileName("c:\")
    Assert.AreEqual "file.txt", Path.GetFileName(":file.txt")
    Assert.AreEqual ".txt", Path.GetFileName(".txt")
End Sub

Public Sub TestGetFileNameWithoutExtension()
    Assert.AreEqual "file", Path.GetFileNameWithoutExtension("file.txt")
    Assert.AreEqual "file", Path.GetFileNameWithoutExtension("c:\file.txt")
    Assert.AreEqual "", Path.GetFileNameWithoutExtension("c:\")
    Assert.AreEqual "file", Path.GetFileNameWithoutExtension("c:file.txt")
    Assert.AreEqual "file", Path.GetFileNameWithoutExtension("file")
    Assert.AreEqual "file", Path.GetFileNameWithoutExtension("c:file")
    Assert.AreEqual "", Path.GetFileNameWithoutExtension("c:.file")
    Assert.AreEqual "", Path.GetFileNameWithoutExtension("")
    Assert.AreEqual "file.txt", Path.GetFileNameWithoutExtension("c:\file.txt.txt")
    Assert.AreEqual "file.txt", Path.GetFileNameWithoutExtension("file.txt.txt")
    Assert.AreEqual "", Path.GetFileNameWithoutExtension("file.txt\")
    Assert.AreEqual "", Path.GetFileNameWithoutExtension(".txt")
End Sub

Public Sub TestGetFullPathOfLegalPaths()
    Assert.AreEqual CurDir & "\file.txt", Path.GetFullPath("file.txt")
    Assert.AreEqual Path.GetPathRoot(CurDir) & "file.txt", Path.GetFullPath("\file.txt")
    Assert.AreEqual "\\server\share", Path.GetFullPath("\\server\share")
    Assert.AreEqual "\\server\share\", Path.GetFullPath("\\server\share\")
    Assert.AreEqual "c:\abc.txt", Path.GetFullPath("c:\abc.txt")
    Assert.AreEqual CurDir & "\c\abc.txt", Path.GetFullPath("c\abc.txt")
    Assert.AreEqual CurDir & "\a\b\a", Path.GetFullPath("a\b...\a")
    Assert.AreEqual CurDir & "\a\...b\a", Path.GetFullPath("a\...b\a")
    Assert.AreEqual "c:\", Path.GetFullPath("c:\")
    Assert.AreEqual "c:\", Path.GetFullPath("c:\\\")
    Assert.AreEqual "\\server\share", Path.GetFullPath("\\\\server\\\share")
    Assert.AreEqual "\\server\share\", Path.GetFullPath("\\\\server\\\share\\\")
    Assert.AreEqual "c:\abc\jkl", Path.GetFullPath("c:\abc\def\ghi\..\..\jkl")
    Assert.AreEqual "c:\jkl", Path.GetFullPath("c:\abc\def\ghi\..\..\..\..\..\jkl")
    Assert.AreEqual Left$(CurDir, InStrRev(CurDir, "\")) & "jkl", Path.GetFullPath("..\jkl")
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
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub TestGetFullPathInvalidUNCShareWithSlash()
    On Error GoTo errTrap
    Call Path.GetFullPath("\\share\")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub TestGetFullPathOfEmptyPath()
    On Error GoTo errTrap
    Call Path.GetFullPath("")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub TestGetFullPathOfSpaces()
    On Error GoTo errTrap
    Call Path.GetFullPath("  ")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type."
End Sub

Public Sub TestGetFullPathMultipleColons()
    On Error GoTo errTrap
    Path.GetFullPath "c:\a:"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
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

