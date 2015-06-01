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
' DirectoryTests
'
Option Explicit
Implements ITestCaseSource
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

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetDirectoryRoot_WithDirectories_ReturnsExpected"
            Test.Use "c:\somepath\", "c:\"
            Test.Use "\\server\share\somefolder\", "\\server\share"
            Test.Use "folder\subfolder", "c:\"
        
        Case "CreateDirectory_WithIllegalCharacters_ThrowsArgumentException"
            Test.Use "c:\a<b"
            Test.Use "c:\a>b"
            Test.Use "c:\a" & Chr$(31) & "b"
            
        Case "GetFiles_WithSearchOption_ReturnsExpected"
            Test.Use SearchOption.TopDirectoryOnly, Array("Top.txt")
            Test.Use SearchOption.AllDirectories, Array("Top.txt", "SubFolder\Sub.txt")
            
        Case "GetDirectories_WithSearchOption_ReturnsExpected"
            Test.Use SearchOption.TopDirectoryOnly, Array("SubFolder")
            Test.Use SearchOption.AllDirectories, Array("SubFolder", "SubFolder\SubSubFolder")
            
    End Select
End Sub

Public Sub Exists_WithExistingDirectory_ReturnsTrue()
    Dim Actual As Boolean
    
    Actual = Directory.Exists(mBuilder.TempFolder)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Exists_WithNonExistingDirectory_ReturnsFalse()
    Dim NonExistingDirectory As String
    NonExistingDirectory = Path.Combine(mBuilder.TempFolder, "NonExisting")
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

Public Sub CreateDirectory_WithIllegalCharacters_ThrowsArgumentException(ByVal IllegalPath As String)
    On Error GoTo errTrap
    
    Directory.CreateDirectory IllegalPath
    
errTrap:
    AssertArgumentException Err
End Sub

Public Sub CreateDirectory_WithColonAfterRoot_ThrowsNotSupportedException()
    On Error GoTo errTrap
    
    Directory.CreateDirectory "c:\abc:d"
    
errTrap:
    AssertNotSupportedException Err
End Sub

Public Sub Delete_WhenCalled_DeletesDirectory()
    Directory.Delete mBuilder.TempFolder
    
    Assert.That Helper.FolderExists(mBuilder.TempFolder), Iz.False
End Sub

Public Sub Delete_WithRecursive_DeletesSubfoldersAndFiles()
    mBuilder.CreateFolder
    mBuilder.CreateFile
    
    Directory.Delete mBuilder.TempFolder, True
    
    Assert.That Helper.FolderExists(mBuilder.TempFolder), Iz.False
End Sub

Public Sub Delete_WithRecursiveAndMultipleLevels_DeletesFolder()
    mBuilder.CreateFolder
    mBuilder.CreateFolder "SubFolder\SubSubFolder"
    mBuilder.CreateFile "SubFolder\SubSubFolder\File.txt"
    
    Directory.Delete mBuilder.TempFolder, True
    
    Assert.That FolderExists(mBuilder.TempFolder), Iz.False
End Sub

Public Sub GetDirectories_WhenCalled_ReturnsDirectories()
    Dim Expected As String
    Expected = mBuilder.CreateFolder
    
    Dim Actual() As String
    Actual = Directory.GetDirectories(mBuilder.TempFolder)
    
    Assert.That Actual, Iz.EqualTo(Array(Expected))
End Sub

Public Sub GetDirectories_WithSearchPattern_ReturnsSelectDirectories()
    Dim Expected As String
    mBuilder.CreateFolder "SubFolder1"
    Expected = mBuilder.CreateFolder("SubFolder2")
    
    Dim Actual() As String
    Actual = Directory.GetDirectories(mBuilder.TempFolder, "*2")
    
    Assert.That Actual, Iz.EqualTo(Array(Expected))
End Sub

Public Sub GetDirectoryRoot_WithDirectories_ReturnsExpected(ByVal DirectoryToUse As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Directory.GetDirectoryRoot(DirectoryToUse)

    Assert.That Actual, Iz.EqualTo(Expected).IgnoreCase
End Sub

Public Sub GetFiles_WhenCalled_ReturnsFiles()
    Dim Expected As String
    Expected = mBuilder.CreateFile
    
    Dim Actual() As String
    Actual = Directory.GetFiles(mBuilder.TempFolder)
    
    Assert.That Actual, Iz.EqualTo(Array(Expected))
End Sub

Public Sub GetFiles_WithSearchPattern_ReturnsSelectFiles()
    Dim Expected As String
    mBuilder.CreateFile "File1.txt"
    Expected = mBuilder.CreateFile("File2.txt")
    
    Dim Actual() As String
    Actual = Directory.GetFiles(mBuilder.TempFolder, "*2.txt")
    
    Assert.That Actual, Iz.EqualTo(Array(Expected))
End Sub

Public Sub GetFiles_WithSearchOption_ReturnsExpected(ByVal Opt As SearchOption, ByRef Expected As Variant)
    mBuilder.CreateFile "Top.txt"
    mBuilder.CreateFolder
    mBuilder.CreateFile "SubFolder\Sub.txt"
    
    Dim Actual() As String
    Actual = Directory.GetFiles(mBuilder.TempFolder, SearchOpt:=Opt)
    
    Assert.That Actual, Iz.EqualTo(CreateFullPaths(Expected))
End Sub

Public Sub GetDirectories_WithSearchOption_ReturnsExpected(ByVal Opt As SearchOption, ByRef Expected As Variant)
    mBuilder.CreateFolder
    mBuilder.CreateFolder "SubFolder\SubSubFolder"
    
    Dim Actual() As String
    Actual = Directory.GetDirectories(mBuilder.TempFolder, SearchOpt:=Opt)
    
    Assert.That Actual, Iz.EqualTo(CreateFullPaths(Expected))
End Sub

Public Sub GetFileSystemEntries_WithFilesAndFolders_ReturnsBothTypes()
    Dim Expected As New Collection
    Expected.Add mBuilder.CreateFolder
    Expected.Add mBuilder.CreateFile
    
    Dim Actual() As String
    Actual = Directory.GetFileSystemEntries(mBuilder.TempFolder)
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetFileSystemEntries_WithSubFoldersAndFiles_ReturnsAll()
    Dim Expected As New Collection
    Expected.Add mBuilder.CreateFolder
    Expected.Add mBuilder.CreateFile
    Expected.Add mBuilder.CreateFolder("SubFolder\SubSubFolder")
    Expected.Add mBuilder.CreateFile("SubFolder\File.txt")
    
    Dim Actual() As String
    Actual = Directory.GetFileSystemEntries(mBuilder.TempFolder, SearchOpt:=SearchOption.AllDirectories)
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub CreationTime_WhenCalled_IsSetAndRead()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    
    Directory.SetCreationTime FileName, #2/2/2015 8:40:00 AM#
    Dim Actual As Date
    Actual = Directory.GetCreationTime(FileName).ToOADate
    
    Assert.That Actual, Iz.EqualTo(#2/2/2015 8:40:00 AM#)
End Sub

Public Sub LastAccessTime_WhenCalled_IsSetAndRead()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    
    Directory.SetLastAccessTime FileName, #2/2/2015 8:40:00 AM#
    Dim Actual As Date
    Actual = Directory.GetLastAccessTime(FileName).ToOADate
    
    Assert.That Actual, Iz.EqualTo(#2/2/2015 8:40:00 AM#)
End Sub

Public Sub LastWriteTime_WhenCalled_IsSetAndRead()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    
    Directory.SetLastWriteTime FileName, #2/2/2015 8:40:00 AM#
    Dim Actual As Date
    Actual = Directory.GetLastWriteTime(FileName).ToOADate
    
    Assert.That Actual, Iz.EqualTo(#2/2/2015 8:40:00 AM#)
End Sub

Public Sub LastWriteTime_WithNonExistingFile_ThrowsFileNotFoundException()
    Dim NonExistingFileName As String
    NonExistingFileName = Path.Combine(mBuilder.TempFolder, "NothingHere.txt")
    
    On Error Resume Next
    Directory.SetLastWriteTime NonExistingFileName, Now
    
    AssertFileNotFoundException Err, NonExistingFileName
End Sub

Public Sub GetParent_WithFile_ReturnsParentDirectory()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    
    Dim Actual As DirectoryInfo
    Set Actual = Directory.GetParent(FileName)
    
    Assert.That Actual.FullName, Iz.EqualTo(mBuilder.TempFolder)
End Sub

Public Sub GetParent_WithSubFolder_ReturnsParentDirectory()
    Dim SubFolder As String
    SubFolder = mBuilder.CreateFolder
    
    Dim Actual As DirectoryInfo
    Set Actual = Directory.GetParent(SubFolder)
    
    Assert.That Actual.FullName, Iz.EqualTo(mBuilder.TempFolder)
End Sub

Public Sub GetParent_WithEmptyPath_ThrowsArgumentException()
    On Error GoTo errTrap
    
    Directory.GetParent ""
    
errTrap:
    AssertArgumentException Err, "Path"
End Sub

Public Sub GetParent_WithRootPath_ReturnsNothing()
    Dim Actual As DirectoryInfo
    
    Set Actual = Directory.GetParent("c:\")
    
    Assert.That Actual, Iz.Nothing
End Sub

Public Sub Move_WhenCalled_MovesContents()
    Dim SourceFolder As String
    SourceFolder = mBuilder.CreateFolder("Source")
    mBuilder.CreateFile "Source\File.txt"
    Dim TargetFolder As String
    TargetFolder = Path.Combine(mBuilder.TempFolder, "Target")
    
    Directory.Move SourceFolder, TargetFolder
    
    Assert.That FolderExists(SourceFolder), Iz.False
    Assert.That FolderExists(TargetFolder), Iz.True
    Assert.That FileExists(Path.Combine(TargetFolder, "File.txt")), Iz.True
End Sub

Public Sub CurrentDirectory_WhenCalled_IsSetAndRead()
    Dim OriginalCurrentDirectory As String
    OriginalCurrentDirectory = CurDir$
    Dim NewFolder As String
    NewFolder = mBuilder.CreateFolder
    
    On Error GoTo errTrap
    Directory.SetCurrentDirectory NewFolder
    Dim Actual As String
    Actual = Directory.GetCurrentDirectory
    On Error GoTo 0
    
    Assert.That Actual, Iz.EqualTo(NewFolder)
errTrap:
    ChDir OriginalCurrentDirectory
End Sub

Private Function CreateFullPaths(ByVal Source As Variant) As Collection
    Set CreateFullPaths = New Collection
    
    Dim Item As Variant
    For Each Item In Source
        CreateFullPaths.Add Path.Combine(mBuilder.TempFolder, CStr(Item))
    Next
End Function


