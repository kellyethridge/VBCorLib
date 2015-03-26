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

Private mBuilder As DirectoryBuilder

Public Sub Setup()
    Set mBuilder = New DirectoryBuilder
End Sub

Public Sub TearDown()
    Set mBuilder = Nothing
End Sub

Public Sub Exists_WithExistingDirectory_ReturnsTrue()
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As Boolean
    Actual = Info.Exists
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Exists_WithNonExistingDirectory_ReturnsFalse()
    Dim NonExistingDirectory As String
    NonExistingDirectory = Path.Combine(mBuilder.TempFolder, "NonExisting")
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(NonExistingDirectory)
    
    Dim Actual As Boolean
    Actual = Info.Exists

    Assert.That Actual, Iz.False
End Sub

Public Sub Create_WhenCalled_CreatesFolder()
    Dim FolderToCreate As String
    FolderToCreate = Path.Combine(Path.GetTempPath, "SubFolder")
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(FolderToCreate)
    
    Info.Create

    Assert.That FolderExists(FolderToCreate), Iz.True
    Assert.That Info.Exists, Iz.True
End Sub

Public Sub Name_WhenCalled_ReturnsTheLastDirectoryName()
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As String
    Actual = Info.Name
    
    Assert.That Actual, Iz.EqualTo(mBuilder.TempSubFolder)
End Sub

Public Sub FullName_WhenCalled_ReturnsEntirePath()
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As String
    Actual = Info.FullName
    
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
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(Folder)
    
    Dim Actual As FileAttributes
    Actual = Info.Attributes And Hidden
    
    Assert.That Actual, Iz.EqualTo(FileAttributes.Hidden)
End Sub

Public Sub Attributes_WhenSet_SetsIt()
    Dim Folder As String
    Folder = mBuilder.CreateFolder
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(Folder)
    
    Info.Attributes = FileAttributes.Hidden
    
    Assert.That GetAttr(Folder), Iz.EqualTo(vbDirectory Or vbHidden)
End Sub

Public Sub CreationTime_WhenCalled_IsSetAndRead()
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.CreateFile)
    
    Info.CreationTime = #2/2/2015 8:40:00 AM#
    Dim Actual As Date
    Actual = Info.CreationTime.ToOADate
    
    Assert.That Actual, Iz.EqualTo(#2/2/2015 8:40:00 AM#)
End Sub

Public Sub LastAccessTime_WhenCalled_IsSetAndRead()
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.CreateFile)
    
    Info.LastAccessTime = #2/2/2015 8:40:00 AM#
    Dim Actual As Date
    Actual = Info.LastAccessTime.ToOADate
    
    Assert.That Actual, Iz.EqualTo(#2/2/2015 8:40:00 AM#)
End Sub

Public Sub LastWriteTime_WhenCalled_IsSetAndRead()
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.CreateFile)
    
    Info.LastWriteTime = #2/2/2015 8:40:00 AM#
    Dim Actual As Date
    Actual = Info.LastWriteTime.ToOADate
    
    Assert.That Actual, Iz.EqualTo(#2/2/2015 8:40:00 AM#)
End Sub

Public Sub Root_WhenCalled_ReturnsFolderRoot()
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As DirectoryInfo
    Set Actual = Info.Root
    
    Assert.That Actual.FullName, Iz.EqualTo("c:\").IgnoreCase
End Sub

Public Sub GetDirectories_WhenCalled_ReturnsTopDirectories()
    Dim Expected As New Collection
    Expected.Add mBuilder.CreateFolder("SubFolder1")
    Expected.Add mBuilder.CreateFolder("SubFolder2")
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As Collection
    Set Actual = GetPathNames(Info.GetDirectories)
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetDirectories_WithSearchPattern_ReturnsMatching()
    Dim Expected As New Collection
    mBuilder.CreateFolder "SubFolder1"
    Expected.Add mBuilder.CreateFolder("SubFolder2")
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As Collection
    Set Actual = GetPathNames(Info.GetDirectories("*2"))
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetDirectories_WithTopDirectoriesSearchOption_ReturnsOnlyTopDirectory()
    Dim Expected As New Collection
    Expected.Add mBuilder.CreateFolder
    mBuilder.CreateFolder "SubFolder\SubSubFolder"
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As Collection
    Set Actual = GetPathNames(Info.GetDirectories(SearchOpt:=SearchOption.TopDirectoryOnly))
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetDirectories_WithAllDirectoriesSearchOption_ReturnsAllFolders()
    Dim Expected As New Collection
    Expected.Add mBuilder.CreateFolder
    Expected.Add mBuilder.CreateFolder("SubFolder\SubSubFolder")
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual As Collection
    Set Actual = GetPathNames(Info.GetDirectories(SearchOpt:=SearchOption.AllDirectories))
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetDirectories_WithNoDirectories_ReturnsEmptyArray()
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mBuilder.TempFolder)
    
    Dim Actual() As DirectoryInfo
    Actual = Info.GetDirectories
    
    Assert.That Actual, Iz.Empty
End Sub

Private Function GetPathNames(ByRef Infos() As DirectoryInfo) As Collection
    Set GetPathNames = New Collection
    
    Dim Info As Variant
    For Each Info In Infos
        GetPathNames.Add Info.FullName
    Next
End Function

'Private info As DirectoryInfo
'
'Public Sub TestFixtureSetup()
'    WriteLine ""
'    WriteLine "***** DIRECTORYINFO *****"
'End Sub
'
'Public Sub Setup()
'    Set info = NewDirectoryInfo(App.Path)
'End Sub
'
'Public Sub TestCreationTimeLocalVsUTC()
'    Dim loc As cDateTime
'    Dim utc As cDateTime
'    Dim Offset As TimeSpan
'
'    Set loc = info.CreationTime
'    Set utc = info.CreationTimeUtc
'    Set Offset = TimeZone.CurrentTimeZone.GetUtcOffset(loc)
'
'    WriteLine cString.Format("Local: {0}", loc)
'    WriteLine cString.Format("UTC: {0}", utc)
'    WriteLine cString.Format("Offset: {0}", Offset)
'
'    Assert.AreEqual Offset.TotalHours, loc.Subtract(utc).TotalHours
'End Sub
'
'Public Sub TestGetDirectories()
'    Dim di() As DirectoryInfo
'    Dim i As Long
'
'    di = info.GetDirectories
'    Assert.IsFalse cArray.IsNull(di), "Should not be a null array."
'
'    WriteLine ""
'    WriteLine "--- VISUAL INSPECTION ---"
'    WriteLine "GetDirectories"
'    WriteLine "Count: " & cArray.GetLength(di)
'    For i = 0 To UBound(di)
'        WriteLine di(i).Name
'    Next i
'End Sub
'
'Public Sub TestEmptyFolder()
'    Dim di() As DirectoryInfo
'    Dim EmptyFolder As DirectoryInfo
'
'    Set EmptyFolder = Cor.NewDirectoryInfo(Path.Combine(info.FullName, "Folder"))
'
'    di = EmptyFolder.GetDirectories
'
'    Assert.IsFalse cArray.IsNull(di)
'    Assert.AreEqual 0, cArray.GetLength(di)
'End Sub
'
'Public Sub TestNoParent()
'    Set info = Cor.NewDirectoryInfo("c:\")
'    Assert.IsNothing info.Parent, "'c:\' should have no parent."
'
'    Set info = Cor.NewDirectoryInfo("\\server\share")
'    Assert.IsNothing info.Parent, "'\\server\share' should have no parent."
'End Sub
'
'Public Sub TestGetDirectoriesNoMatch()
'    Dim di() As DirectoryInfo
'
'    di = info.GetDirectories("abc123")
'    Assert.IsFalse cArray.IsNull(di)
'    Assert.AreEqual 0, cArray.GetLength(di)
'End Sub
'
'Public Sub TestGetFilesNoMatch()
'    Dim fi() As FileInfo
'
'    fi = info.GetFiles("abc123")
'    Assert.IsFalse cArray.IsNull(fi)
'    Assert.AreEqual 0, cArray.GetLength(fi)
'End Sub
