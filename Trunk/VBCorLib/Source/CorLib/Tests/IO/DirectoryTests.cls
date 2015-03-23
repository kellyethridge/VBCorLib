VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectoryTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDirectory
'
Option Explicit
Implements ITestCaseSource

Private Const TempSubFolder As String = "VBCorlib.Directory.Tests"

Private mTempFolder As String


Public Sub Setup()
    mTempFolder = Path.Combine(Path.GetTempPath, TempSubFolder)
    
    If Not Directory.Exists(mTempFolder) Then
        Directory.CreateDirectory mTempFolder
    End If
End Sub

Public Sub Teardown()
    If Directory.Exists(mTempFolder) Then
        Directory.Delete mTempFolder, True
    End If
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetDirectoryRoot_WithDirectories_ReturnsExpected"
            Test.Use "c:\somepath\", "c:\"
            Test.Use "\\server\share\somefolder\", "\\server\share"
            Test.Use "folder\subfolder", "c:\"
        
    End Select
End Sub

Public Sub Exists_WithExistingDirectory_ReturnsTrue()
    Dim Actual As Boolean
    
    Actual = Directory.Exists(mTempFolder)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Exists_WithNonExistingDirectory_ReturnsFalse()
    Dim NonExistingDirectory As String
    NonExistingDirectory = Path.Combine(mTempFolder, "NonExisting")
    Dim Actual As Boolean
    
    Actual = Directory.Exists(NonExistingDirectory)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub CreateDirectory_WhenCalled_CreatesFolder()
    Const SubFolder As String = "VBCorLib.Directory.CreateDirectory.Test"
    On Error GoTo errTrap
    Dim FolderToCreate As String
    FolderToCreate = Path.Combine(Path.GetTempPath, SubFolder)
    
    Directory.CreateDirectory FolderToCreate
    
    Dim Actual As Boolean
    Actual = Helper.FolderExists(FolderToCreate)
    Assert.That Actual, Iz.True
    
errTrap:
    Helper.DeleteFolder FolderToCreate
End Sub

Public Sub CreateDirectory_WhenCalled_ReturnsDirectoryInfo()
    Const SubFolder As String = "VBCorLib.Directory.CreateDirectory.Test"
    On Error GoTo errTrap
    Dim FolderToCreate As String
    FolderToCreate = Path.Combine(Path.GetTempPath, SubFolder)
    
    Dim Actual As DirectoryInfo
    Set Actual = Directory.CreateDirectory(FolderToCreate)

    Assert.That Actual.Exists, Iz.True

errTrap:
    Helper.DeleteFolder FolderToCreate
End Sub

Public Sub CreateDirectory_WithEmptyPath_ThrowsArgumentException()
    On Error GoTo errTrap
    
    Directory.CreateDirectory ""
    
errTrap:
    AssertArgumentException Err, "Path"
End Sub

Public Sub Delete_WhenCalled_DeletesDirectory()
    Directory.Delete mTempFolder
    
    Assert.That Helper.FolderExists(mTempFolder), Iz.False
End Sub

Public Sub Delete_WithRecursive_DeletesSubfoldersAndFiles()
    Helper.CreateFolder Path.Combine(mTempFolder, "SubFolder")
    Helper.CreateFile Path.Combine(mTempFolder, "File.txt")
    
    Directory.Delete mTempFolder, True
    
    Assert.That Helper.FolderExists(mTempFolder), Iz.False
End Sub

Public Sub GetDirectories_WhenCalled_ReturnsDirectories()
    Dim SubFolder As String
    SubFolder = Path.Combine(mTempFolder, "GetDirectories")
    Helper.CreateFolder SubFolder
    
    Dim Folders() As String
    Folders = Directory.GetDirectories(mTempFolder)
    
    Assert.That Folders, Iz.EqualTo(Array(SubFolder))
End Sub

Public Sub GetDirectories_WithSearchPattern_ReturnsSelectDirectories()
    Dim SubFolder1 As String: SubFolder1 = Path.Combine(mTempFolder, "GetDirectories1")
    Dim SubFolder2 As String: SubFolder2 = Path.Combine(mTempFolder, "GetDirectories2")
    Helper.CreateFolder SubFolder1
    Helper.CreateFolder SubFolder2
    
    Dim Folders() As String
    Folders = Directory.GetDirectories(mTempFolder, "*2")
    
    Assert.That Folders, Iz.EqualTo(Array(SubFolder2))
End Sub

Public Sub GetDirectoryRoot_WithDirectories_ReturnsExpected(ByVal DirectoryToUse As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Directory.GetDirectoryRoot(DirectoryToUse)

    Assert.That Actual, Iz.EqualTo(Expected).IgnoreCase
End Sub

Public Sub GetFiles_WhenCalled_ReturnsFiles()
    Dim FilePath As String
    FilePath = Path.Combine(mTempFolder, "AFile.txt")
    Helper.CreateFile FilePath
    
    Dim Actual() As String
    Actual = Directory.GetFiles(mTempFolder)
    
    Assert.That Actual, Iz.EqualTo(Array(FilePath))
End Sub

Public Sub GetFiles_WithSearchPattern_ReturnsSelectFiles()
    Dim FilePath1 As String: FilePath1 = Path.Combine(mTempFolder, "GetFiles1.txt")
    Dim FilePath2 As String: FilePath2 = Path.Combine(mTempFolder, "GetFiles2.txt")
    Helper.CreateFile FilePath1
    Helper.CreateFile FilePath2
    
    Dim Actual() As String
    Actual = Directory.GetFiles(mTempFolder, "*2.txt")
    
    Assert.That Actual, Iz.EqualTo(Array(FilePath2))
End Sub

'Public Sub TestGetFileSystemEntries()
'    Dim col As New Collection
'    Dim s As String
'    s = Dir$(App.Path & "\*", vbDirectory)
'    Do While Len(s) > 0
'        If s <> "." And s <> ".." Then
'            col.Add App.Path & "\" & s
'        End If
'        s = Dir$
'    Loop
'
'    Dim files() As String
'    Dim i As Long
'
'    files = Directory.GetFileSystemEntries(App.Path)
'    Assert.AreEqual col.Count, cArray.GetLength(files)
'
'    For i = 0 To UBound(files)
''        Assert.Contains files(i), col
'        Assert.That col, Has.Member(files(i))
'    Next i
'End Sub
