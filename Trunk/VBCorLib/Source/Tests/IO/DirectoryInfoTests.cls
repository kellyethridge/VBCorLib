VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectoryInfoTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDirectoryInfo
'
Option Explicit

Private mBuilder As IOHelper

Public Sub Setup()
    Set mBuilder = New IOHelper
End Sub

Public Sub Teardown()
    Set mBuilder = Nothing
End Sub

Public Sub Exists_WithExistingDirectory_ReturnsTrue()
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As Boolean
    Actual = info.Exists
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Exists_WithNonExistingDirectory_ReturnsFalse()
    Dim NonExistingDirectory As String
    NonExistingDirectory = Path.Combine(mBuilder.TempFolder, "NonExisting")
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(NonExistingDirectory)
    
    Dim Actual As Boolean
    Actual = info.Exists

    Assert.That Actual, Iz.False
End Sub

Public Sub Create_WhenCalled_CreatesFolder()
    Dim FolderToCreate As String
    FolderToCreate = Path.Combine(Path.GetTempPath, "SubFolder")
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(FolderToCreate)
    
    info.Create

    Assert.That FolderExists(FolderToCreate), Iz.True
    Assert.That info.Exists, Iz.True
End Sub

Public Sub Name_WhenCalled_ReturnsTheLastDirectoryName()
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As String
    Actual = info.Name
    
    Assert.That Actual, Iz.EqualTo(mBuilder.TempSubFolder)
End Sub

Public Sub FullName_WhenCalled_ReturnsEntirePath()
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As String
    Actual = info.FullName
    
    Assert.That Actual, Iz.EqualTo(mBuilder.TempFolder)
End Sub

Public Sub Parent_WhenCalled_ReturnsParentInfo()
    Dim SubFolder As DirectoryInfo
    Set SubFolder = Cor.NewDirectoryInfo(mBuilder.CreateFolder)
    Dim ParentFolder As DirectoryInfo
    Set ParentFolder = SubFolder.Parent
    
    Dim Actual As String
    Actual = ParentFolder.FullName
    
    Assert.That Actual, Iz.EqualTo(mBuilder.TempFolder)
End Sub

Public Sub Parent_WithRootFolder_ReturnsNothing()
    Dim RootFolder As DirectoryInfo
    Set RootFolder = Cor.NewDirectoryInfo("c:\")
    
    Dim Actual As DirectoryInfo
    Set Actual = RootFolder.Parent
    
    Assert.That Actual, Iz.Nothing
End Sub

Public Sub Attributes_WhenRead_ReadsIt()
    Dim Folder As String
    Folder = mBuilder.CreateFolder
    SetAttr Folder, vbHidden
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(Folder)
    
    Dim Actual As FileAttributes
    Actual = info.Attributes And Hidden
    
    Assert.That Actual, Iz.EqualTo(FileAttributes.Hidden)
End Sub

Public Sub Attributes_WhenSet_SetsIt()
    Dim Folder As String
    Folder = mBuilder.CreateFolder
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(Folder)
    
    info.Attributes = FileAttributes.Hidden
    
    Assert.That GetAttr(Folder), Iz.EqualTo(vbDirectory Or vbHidden)
End Sub

Public Sub CreationTime_WhenCalled_IsSetAndRead()
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.CreateFile)
    
    info.CreationTime = #2/2/2015 8:40:00 AM#
    Dim Actual As Date
    Actual = info.CreationTime.ToOADate
    
    Assert.That Actual, Iz.EqualTo(#2/2/2015 8:40:00 AM#)
End Sub

Public Sub LastAccessTime_WhenCalled_IsSetAndRead()
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.CreateFile)
    
    info.LastAccessTime = #2/2/2015 8:40:00 AM#
    Dim Actual As Date
    Actual = info.LastAccessTime.ToOADate
    
    Assert.That Actual, Iz.EqualTo(#2/2/2015 8:40:00 AM#)
End Sub

Public Sub LastWriteTime_WhenCalled_IsSetAndRead()
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.CreateFile)
    
    info.LastWriteTime = #2/2/2015 8:40:00 AM#
    Dim Actual As Date
    Actual = info.LastWriteTime.ToOADate
    
    Assert.That Actual, Iz.EqualTo(#2/2/2015 8:40:00 AM#)
End Sub

Public Sub Root_WhenCalled_ReturnsFolderRoot()
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As DirectoryInfo
    Set Actual = info.Root
    
    Assert.That Actual.FullName, Iz.EqualTo("c:\").IgnoreCase
End Sub

Public Sub GetDirectories_WhenCalled_ReturnsTopDirectories()
    Dim Expected As New Collection
    Expected.Add mBuilder.CreateFolder("SubFolder1")
    Expected.Add mBuilder.CreateFolder("SubFolder2")
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As Collection
    Set Actual = GetPathNames(info.GetDirectories)
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetDirectories_WithSearchPattern_ReturnsMatching()
    Dim Expected As New Collection
    mBuilder.CreateFolder "SubFolder1"
    Expected.Add mBuilder.CreateFolder("SubFolder2")
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As Collection
    Set Actual = GetPathNames(info.GetDirectories("*2"))
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetDirectories_WithNoMatchingSearchPattern_ReturnsEmpty()
    mBuilder.CreateFolder
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual() As DirectoryInfo
    Actual = info.GetDirectories("NoMatch")
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub GetDirectories_WithTopDirectoriesSearchOption_ReturnsOnlyTopDirectory()
    Dim Expected As New Collection
    Expected.Add mBuilder.CreateFolder
    mBuilder.CreateFolder "SubFolder\SubSubFolder"
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As Collection
    Set Actual = GetPathNames(info.GetDirectories(SearchOpt:=SearchOption.TopDirectoryOnly))
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetDirectories_WithAllDirectoriesSearchOption_ReturnsAllFolders()
    Dim Expected As New Collection
    Expected.Add mBuilder.CreateFolder
    Expected.Add mBuilder.CreateFolder("SubFolder\SubSubFolder")
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As Collection
    Set Actual = GetPathNames(info.GetDirectories(SearchOpt:=SearchOption.AllDirectories))
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetDirectories_WithNoDirectories_ReturnsEmptyArray()
    Dim info As DirectoryInfo
    Set info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual() As DirectoryInfo
    Actual = info.GetDirectories
    
    Assert.That Actual, Iz.Empty
End Sub

Private Function GetPathNames(ByRef Infos() As DirectoryInfo) As Collection
    Set GetPathNames = New Collection
    
    Dim info As Variant
    For Each info In Infos
        GetPathNames.Add info.FullName
    Next
End Function
