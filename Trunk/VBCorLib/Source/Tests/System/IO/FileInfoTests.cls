VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileInfoTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private mBuilder As IOHelper


Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
End Sub

Public Sub Setup()
    Set mBuilder = New IOHelper
End Sub

Public Sub Teardown()
    Set mBuilder = Nothing
End Sub

Public Sub Name_WhenCalled_RetursFileNamePortion()
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo("c:\windows\file.txt")
    
    Dim Actual As String
    Actual = Info.Name
    
    Assert.That Actual, Iz.EqualTo("file.txt")
End Sub

Public Sub FullName_WithFullPath_ReturnsFullPath()
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo("c:\windows\file.txt")
    
    Dim Actual As String
    Actual = Info.FullName
    
    Assert.That Actual, Iz.EqualTo("c:\windows\file.txt")
End Sub

Public Sub FullName_WithFileName_ReturnsFullPath()
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo("file.txt")
    
    Dim Actual As String
    Actual = Info.FullName
    
    Dim Expected As String
    Expected = Path.Combine(CurDir, "file.txt")
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Attributes_WhenSet_AreSetToFile()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abc")
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    Info.Attributes = FileAttributes.Hidden
    
    Dim Actual As FileAttributes
    Actual = Info.Attributes
    Assert.That Actual, Iz.EqualTo(FileAttributes.Hidden)
End Sub

Public Sub Attributes_WhenCalled_ReturnsCurrentAttributes()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abc")
    SetAttr FileName, vbNormal
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    Dim Actual As FileAttributes
    Actual = Info.Attributes
    
    Assert.That Actual, Iz.EqualTo(FileAttributes.Normal)
End Sub

Public Sub Attributes_WhenCalledWithNonExistingFile_ThrowsFileNotFoundException()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "Missing.txt")
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    On Error Resume Next
    Info.Attributes = Normal
    
    AssertFileNotFoundException Err, FileName
End Sub

Public Sub CreationTime_WhenCalled_SetsAndReturnsNewTime()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    Info.CreationTime = #1/1/2001 5:00:00 AM#
    Dim Actual As Date
    Actual = Info.CreationTime.ToOADate

    Assert.That Actual, Iz.EqualTo(#1/1/2001 5:00:00 AM#)
End Sub

Public Sub LastAccessTime_WhenCalled_SetsAndReturnsNewTime()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    Info.LastAccessTime = #1/1/2001 5:00:00 AM#
    Dim Actual As Date
    Actual = Info.LastAccessTime.ToOADate

    Assert.That Actual, Iz.EqualTo(#1/1/2001 5:00:00 AM#)
End Sub

Public Sub LastWriteTime_WhenCalled_SetsAndReturnsNewTime()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    Info.LastWriteTime = #1/1/2001 5:00:00 AM#
    Dim Actual As Date
    Actual = Info.LastWriteTime.ToOADate

    Assert.That Actual, Iz.EqualTo(#1/1/2001 5:00:00 AM#)
End Sub

Public Sub DirectoryName_WhenCalled_ReturnsDirectoryPortionOfPath()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    Dim Actual As String
    Actual = Info.DirectoryName
    
    Dim Expected As String
    Expected = mBuilder.TempFolder
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Exists_WhenCalledWithExistingFile_ReturnsTrue()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    Dim Actual As Boolean
    Actual = Info.Exists
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Exists_WhenCalledWithNonExisitngFile_ReturnsFalse()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "Missing.txt")
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    Dim Actual As Boolean
    Actual = Info.Exists
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Extension_WhenCalled_ReturnsFileExtension()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    Dim Actual As String
    Actual = Info.Extension
    
    Assert.That Actual, Iz.EqualTo(".txt")
End Sub

Public Sub Length_WhenCalled_ReturnsFileLength()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abc")
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)

    Dim Actual As Long
    Actual = Info.Length
    
    Assert.That Actual, Iz.EqualTo(3)
End Sub

Public Sub Delete_WhenCalled_DeletesFile()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abc")
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    
    Info.Delete
    
    Assert.That FileExists(FileName), Iz.False
End Sub

Public Sub Delete_WithNonExistingFile_DoesNotThrowException()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "Missing.txt")
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)

    On Error Resume Next
    Info.Delete
    
    AssertNoException Err
End Sub

Public Sub MoveTo_WhenCalled_MovesFile()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abc")
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    Dim DestName As String
    DestName = Path.Combine(mBuilder.TempFolder, "NewFile.txt")

    Info.MoveTo DestName
    
    Assert.That FileExists(DestName), Iz.True
    Assert.That FileExists(FileName), Iz.False
End Sub

Public Sub MoveTo_WhenCompleted_UpdatesFileInfo()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abc")
    Dim Info As FileInfo
    Set Info = Cor.NewFileInfo(FileName)
    Dim DestName As String
    DestName = Path.Combine(mBuilder.TempFolder, "NewFile.txt")

    Info.MoveTo DestName
    
    Assert.That Info.Name, Iz.EqualTo("NewFile.txt")
    Assert.That Info.FullName, Iz.EqualTo(DestName)
End Sub













