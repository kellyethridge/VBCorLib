VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDirectory
'
Option Explicit

Public Sub TestExists()
    Assert.IsTrue Directory.Exists(App.Path & "\Dummy Test Files")
    Assert.IsFalse Directory.Exists(App.Path & "\TestDirectory.cls")
End Sub

Public Sub TestCreateDeleteDirectory()
    Dim f As String
    
    f = App.Path & "\dummy test files\Sub1\Sub2"
    If Directory.Exists(App.Path & "\dummy test files\Sub1") Then Directory.Delete App.Path & "\dummy test files\Sub1", True
    
    Directory.CreateDirectory f
    Assert.IsTrue Directory.Exists(f)
    Directory.Delete App.Path & "\dummy test files\sub1", True
    Assert.IsFalse Directory.Exists(f)
End Sub

Public Sub TestGetDirectories()
    Dim Dirs As New Collection
    Dim d As String
    
    d = Dir$(App.Path & "\*", vbDirectory)
    Do While Len(d) > 0
        If Left$(d, 1) <> "." And ((GetAttr(App.Path & "\" & d) And vbDirectory) = vbDirectory) Then
            Dirs.Add App.Path & "\" & d
        End If
        d = Dir$
    Loop
    
    Dim s() As String
    s = Directory.GetDirectories(App.Path)
    Assert.AreEqual Dirs.Count, cArray.GetLength(s)
    
    Dim i As Long
    For i = 0 To UBound(s)
'        Assert.Contains s(i), Dirs
        Assert.That Dirs, Has.Member(s(i))
    Next i
End Sub

Public Sub TestGetDirectoriesWithPattern()
    Dim s() As String
    
    s = Directory.GetDirectories(App.Path, "fol*")
    Assert.AreEqual 1, cArray.GetLength(s)
    Assert.AreEqual App.Path & "\Folder", s(0)
End Sub

Public Sub TestGetDirectoryRoot()
    Assert.That Directory.GetDirectoryRoot("c:\somepath\"), Iz.EqualTo("c:\").IgnoreCase
    Assert.That Directory.GetDirectoryRoot("\\server\share\some more folders\"), Iz.EqualTo("\\server\share").IgnoreCase
    Assert.That Directory.GetDirectoryRoot("something\more somthing"), Iz.EqualTo("c:\").IgnoreCase
End Sub

Public Sub TestGetFiles()
    Dim col As New Collection
    Dim s As String
    s = Dir$(App.Path & "\*")
    Do While Len(s) > 0
        If (GetAttr(App.Path & "\" & s) And vbDirectory) = 0 Then
            col.Add App.Path & "\" & s
        End If
        s = Dir$
    Loop
    
    Dim files() As String
    Dim i As Long
    
    files = Directory.GetFiles(App.Path)
    Assert.AreEqual col.Count, cArray.GetLength(files)
    
    For i = 0 To UBound(files)
'        Assert.Contains files(i), col
        Assert.That col, Has.Member(files(i))
    Next i
End Sub

Public Sub TestGetFileSystemEntries()
    Dim col As New Collection
    Dim s As String
    s = Dir$(App.Path & "\*", vbDirectory)
    Do While Len(s) > 0
        If s <> "." And s <> ".." Then
            col.Add App.Path & "\" & s
        End If
        s = Dir$
    Loop
    
    Dim files() As String
    Dim i As Long
    
    files = Directory.GetFileSystemEntries(App.Path)
    Assert.AreEqual col.Count, cArray.GetLength(files)
    
    For i = 0 To UBound(files)
'        Assert.Contains files(i), col
        Assert.That col, Has.Member(files(i))
    Next i
End Sub
