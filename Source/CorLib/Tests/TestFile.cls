VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFile
'
Option Explicit

Public Sub TestExists()
    Dim f As String
    
    f = App.Path & "\dummy test files\Exists.txt"
    Assert.IsTrue File.Exists(f), "Exists.txt should exist."
    Assert.IsFalse File.Exists("missingfile.txt"), "File.Exists should return false."
End Sub

Public Sub TestDirectory()
    Dim f As String
    
    f = App.Path
    Assert.IsFalse File.Exists(f), "A directory should never return true."
End Sub

Public Sub TestCopySourceNotFound()
    On Error GoTo errTrap
    File.Copy App.Path & "\dummy test files\notfound.txt", "\dummy test files\cannotbe.txt"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."

errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is FileNotFoundException, "Wrong exception type caught."
End Sub

Public Sub TestCopyNoOverwrite()
    On Error GoTo errTrap
    File.Copy App.Path & "\dummy test files\source.txt", App.Path & "\dummy test files\Exists.txt"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is IOException, "Wrong exception type caught."
End Sub

Public Sub TestCopyOverExisting()
    TestExists
    
    On Error GoTo errTrap
    File.Copy App.Path & "\dummy test files\Source.txt", App.Path & "\dummy test files\Exists.txt", True
    Exit Sub
    
errTrap:
    Assert.Fail "No exception should be thrown."
End Sub

Public Sub TestCopyOverFolderNoOverwrite()
    On Error GoTo errTrap
    File.Copy App.Path & "\dummy test files\Source.txt", "c:\ATI"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is IOException, "Wrong exception type caught."
End Sub

Public Sub TestGetCreationTime()
    WriteLine vbCrLf & "--- VISUAL INSPECTION ---"
    WriteLine "TestGetCreationTime: " & File.GetCreationTime(App.Path & "\dummy test files\source.txt").ToString
End Sub

Public Sub TestSetGetFileTime()
    Dim d As Date
    Dim f As String
    
    f = App.Path & "\dummy test files\source.txt"
    d = Now
    
    File.SetCreationTime f, d
    File.SetLastAccessTime f, d
    File.SetLastWriteTime f, d

    Assert.AreEqual d, File.GetCreationTime(f).ToOADate, "Wrong Creation time."
    Assert.AreEqual d, File.GetLastAccessTime(f).ToOADate, "Wrong LastAccess time."
    Assert.AreEqual d, File.GetLastWriteTime(f).ToOADate, "Wrong LastWrite time."
End Sub

Public Sub TestSetFileTimeUtc()
    Dim d As Date
    Dim f As String
    
    f = App.Path & "\dummy test files\source.txt"
    d = Now
    
    File.SetCreationTimeutc f, d
    File.SetLastAccessTimeutc f, d
    File.SetLastWriteTimeutc f, d

    Assert.AreEqual d, File.GetCreationTimeutc(f).ToOADate, "Wrong Creation time."
    Assert.AreEqual d, File.GetLastAccessTimeutc(f).ToOADate, "Wrong LastAccess time."
    Assert.AreEqual d, File.GetLastWriteTimeutc(f).ToOADate, "Wrong LastWrite time."
End Sub

Public Sub TestDelete()
    Dim f As String
    
    f = App.Path & "\dummy test files\deleteme.txt"
    If Not File.Exists(f) Then
        Open f For Binary As #1
        Close #1
    End If
    
    Assert.IsTrue File.Exists(f), "The deletable file does not exist."
    File.Delete f
    Assert.IsFalse File.Exists(f), "The deletable file should be deleted."
End Sub

Public Sub TestReadAllLines()
    Dim s() As String
    
    s = File.ReadAllLines(Path.Combine(App.Path, "Dummy Test Files\HebrewCalendar_IsLeapYear.txt"))
    Assert.AreEqual 100, cArray.GetLength(s), "Should contain 100 lines."
    Assert.AreEqual "False", s(UBound(s))
End Sub

Public Sub TestReadAllBytes()
    Dim b() As Byte
    
    b = File.ReadAllBytes(Path.Combine(App.Path, "Dummy Test Files\100bytes.txt"))
    Assert.AreEqual 100, cArray.GetLength(b), "Should contain 100 bytes."
    Assert.AreEqual 48, b(99), "Last byte should be 48 '0'."
End Sub

Public Sub TestWriteAllLines()
    Dim s() As String
    
    s = NewStrings("One", "Two", "Three", "Four", "", "Six", "")
    File.WriteAllLines Path.Combine(App.Path, "Dummy Test Files\TestWriteAllLines.txt"), s
    
    Erase s
    s = File.ReadAllLines(Path.Combine(App.Path, "Dummy Test Files\TestWriteAllLines.txt"))
    Assert.AreEqual 7, cArray.GetLength(s)
    Assert.AreEqual "Four", s(3)
    Assert.AreEqual "", s(4)
    Assert.AreEqual "Six", s(5)
    Assert.AreEqual "", s(6)
End Sub