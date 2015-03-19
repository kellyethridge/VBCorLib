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

Private Const TempSubFolder As String = "VBCorlib.Directory.Tests"

Private mTempFolder As String

Public Sub Setup()
    mTempFolder = Path.Combine(Path.GetTempPath, TempSubFolder)
    
    If Not Helper.FolderExists(mTempFolder) Then
        MkDir mTempFolder
    End If
End Sub

Public Sub Teardown()
    Helper.DeleteFolder mTempFolder
End Sub

Public Sub Exists_WithExistingDirectory_ReturnsTrue()
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(mTempFolder)
    
    Dim Actual As Boolean
    Actual = Info.Exists
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Exists_WithNonExistingDirectory_ReturnsFalse()
    Dim NonExistingDirectory As String
    NonExistingDirectory = Path.Combine(mTempFolder, "NonExisting")
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(NonExistingDirectory)
    
    Dim Actual As Boolean
    Actual = Info.Exists

    Assert.That Actual, Iz.False
End Sub

Public Sub Create_WhenCalled_CreatesFolder()
    Const SubFolder As String = "VBCorLib.DirectoryInfo.Create.Test"
    On Error GoTo errTrap
    Dim FolderToCreate As String
    FolderToCreate = Path.Combine(Path.GetTempPath, SubFolder)
    Dim Info As DirectoryInfo
    Set Info = Cor.NewDirectoryInfo(FolderToCreate)
    
    Info.Create

    Assert.That Helper.FolderExists(FolderToCreate), Iz.True
    Assert.That Info.Exists, Iz.True
    Assert.That Info.Name, Iz.EqualTo(SubFolder)
    Assert.That Info.FullName, Iz.EqualTo(FolderToCreate)

errTrap:
    Helper.DeleteFolder FolderToCreate
End Sub


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
'Public Sub TestAttributes()
'    Assert.AreEqual 16, info.Attributes And FileAttributes.DirectoryAttr
'End Sub
'
'Public Sub TestCreationTime()
'    WriteLine ""
'    WriteLine "--- VISUAL INSPECTION ---"
'    WriteLine "CreationTime: " & info.CreationTime.ToString
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
'Public Sub TestName()
'    Assert.AreEqual "Tests", info.Name
'End Sub
'
'Public Sub TestParentName()
'    Dim P As DirectoryInfo
'
'    Set P = info.Parent
'    Assert.IsNotNothing P
'    Assert.AreEqual "VBCorLibGroup", P.Name
'End Sub
'
'Public Sub TestRoot()
'    Dim r As DirectoryInfo
'
'    Set r = info.Root
'    Assert.IsNotNothing r, "No root information in path."
'    Assert.AreEqual "c:\", LCase$(r.Name)
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
