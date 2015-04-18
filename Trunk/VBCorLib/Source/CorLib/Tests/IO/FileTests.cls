VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' FileTests
'
Option Explicit
Implements ITestCaseSource

Private mBuilder As IOHelper

Public Sub Setup()
    Set mBuilder = New IOHelper
End Sub

Public Sub Teardown()
    Set mBuilder = Nothing
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Exists_WithFileName_ReturnsExpected"
            Test.Use "NonExistingFile.txt", False
            Test.Use "File.txt", True
            Test.Use "AFolder", False
            Test.Use "AFolder\NonExistingFile.txt", False
            Test.Use "AFolder\SubFile.txt", True
            
        Case "Exists_WithParentNavigation_ReturnsExpected"
            Test.Use "File.txt", False
            Test.Use "..\File.txt", True
        
        Case "Delete_WithInvalidPath_ThrowsArgumentException"
            Test.Use ""
            Test.Use "       "
            Test.Use "c:\<\File.txt"
            
        Case "Copy_WithInvalidFileName_ThrowsArgumentException"
            Test.Use "c:\file.txt", "", "DestinationFileName"
            Test.Use "", "c:\file.txt", "SourceFileName"
            Test.Use "    ", "c:\file.txt", "SourceFileName"
            Test.Use "c:\file.txt", "    ", "DestinationFileName"
            Test.Use "c:\>\file.txt", "c:\file2.txt", "SourceFileName"
            Test.Use "c:\file.txt", "c:\<file.txt", "DestinationFileName"
            
    End Select
End Sub

Public Sub Exists_WithFileName_ReturnsExpected(ByVal FileName As String, ByVal Expected As Boolean)
    FileName = Path.Combine(mBuilder.TempFolder, FileName)
    mBuilder.CreateFile "File.txt"
    mBuilder.CreateFolder "AFolder"
    mBuilder.CreateFile "AFolder\SubFile.txt"
    
    Dim Actual As Boolean
    Actual = File.Exists(FileName)

    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Exists_WithParentNavigation_ReturnsExpected(ByVal FileName As String, ByVal Expected As Boolean)
    Dim SubFolder As String
    SubFolder = mBuilder.CreateFolder("AFolder")
    mBuilder.CreateFile "File.txt"
    mBuilder.CreateFile "AFolder\SubFile.txt"
    Directory.SetCurrentDirectory SubFolder
        
    Dim Actual As Boolean
    Actual = File.Exists(FileName)

    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Copy_WithInvalidFileName_ThrowsArgumentException(ByVal SourceFileName As String, ByVal DestinationFileName As String, ByVal ParamName As String)
    On Error Resume Next
    
    File.Copy SourceFileName, DestinationFileName
    
    AssertArgumentException Err, ParamName
End Sub

Public Sub Copy_WithMissingFile_ThrowsFileNotFoundException()
    Dim MissingFileName As String
    MissingFileName = Path.Combine(mBuilder.TempFolder, "MissingFile.txt")
    mBuilder.CreateFile
    
    On Error Resume Next
    File.Copy MissingFileName, "NoFile2.txt"
    
    AssertFileNotFoundException Err
End Sub

Public Sub Copy_WithExistingFile_CopiesFile()
    Dim SourceFileName  As String
    Dim NewFileName     As String
    SourceFileName = mBuilder.CreateFile
    NewFileName = Path.Combine(mBuilder.TempFolder, "File (Copy).txt")
    
    File.Copy SourceFileName, NewFileName
    
    Assert.That File.Exists(NewFileName), Iz.True
    Assert.That File.Exists(SourceFileName), Iz.True
End Sub

Public Sub Copy_WithExistingDestinationFile_ThrowsIOException()
    Dim SourceFileName As String
    Dim DestinationFileName As String
    SourceFileName = mBuilder.CreateFile
    DestinationFileName = mBuilder.CreateFile("Existing.txt")
    
    On Error Resume Next
    File.Copy SourceFileName, DestinationFileName
    
    AssertIOException Err
End Sub

Public Sub Copy_WhenOverridingExisting_CopiesFile()
    Dim SourceFileName  As String
    Dim NewFileName     As String
    SourceFileName = mBuilder.CreateFile(Data:="Copied")
    NewFileName = mBuilder.CreateFile("Existing.txt")
    
    File.Copy SourceFileName, NewFileName, True

    Dim Actual As String
    Actual = Helper.ReadFile(NewFileName)
    Assert.That Actual, Iz.EqualTo("Copied")
End Sub

Public Sub Copy_WithDestinationNameAsFolder_ThrowsIOException()
    Dim SourceFileName As String
    Dim TargetFolder As String
    SourceFileName = mBuilder.CreateFile
    TargetFolder = mBuilder.CreateFolder
    
    On Error Resume Next
    File.Copy SourceFileName, TargetFolder
    
    AssertIOException Err
End Sub

Public Sub CreationTime_WhenCalled_SetsAndReturnsNewTime()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    
    File.SetCreationTime FileName, #1/1/2001 5:00:00 AM#
    Dim Actual As Date
    Actual = File.GetCreationTime(FileName).ToOADate

    Assert.That Actual, Iz.EqualTo(#1/1/2001 5:00:00 AM#)
End Sub

Public Sub LastAccessTime_WhenCalled_SetsAndReturnsNewTime()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    
    File.SetLastAccessTime FileName, #1/1/2001 5:00:00 AM#
    Dim Actual As Date
    Actual = File.GetLastAccessTime(FileName).ToOADate

    Assert.That Actual, Iz.EqualTo(#1/1/2001 5:00:00 AM#)
End Sub

Public Sub LastWriteTime_WhenCalled_SetsAndReturnsNewTime()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    
    File.SetLastWriteTime FileName, #1/1/2001 5:00:00 AM#
    Dim Actual As Date
    Actual = File.GetLastWriteTime(FileName).ToOADate

    Assert.That Actual, Iz.EqualTo(#1/1/2001 5:00:00 AM#)
End Sub

Public Sub Delete_WithExistingFiled_DeletesFile()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    
    File.Delete FileName
    
    Assert.That Helper.FileExists(FileName), Iz.False
End Sub

Public Sub Delete_WithInvalidPath_ThrowsArgumentException(ByVal InvalidPath As String)
    On Error Resume Next
    
    File.Delete InvalidPath
    
    AssertArgumentException Err, "Path"
End Sub

'Public Sub TestReadAllLines()
'    Dim s() As String
'
'    s = File.ReadAllLines(Path.Combine(App.Path, "Dummy Test Files\HebrewCalendar_IsLeapYear.txt"))
'    Assert.AreEqual 100, cArray.GetLength(s), "Should contain 100 lines."
'    Assert.AreEqual "False", s(UBound(s))
'End Sub
'
'Public Sub TestReadAllBytes()
'    Dim b() As Byte
'
'    b = File.ReadAllBytes(Path.Combine(App.Path, "Dummy Test Files\100bytes.txt"))
'    Assert.AreEqual 100, cArray.GetLength(b), "Should contain 100 bytes."
'    Assert.AreEqual 48, b(99), "Last byte should be 48 '0'."
'End Sub
'
'Public Sub TestWriteAllLines()
'    Dim s() As String
'
'    s = NewStrings("One", "Two", "Three", "Four", "", "Six", "")
'    File.WriteAllLines Path.Combine(App.Path, "Dummy Test Files\TestWriteAllLines.txt"), s
'
'    Erase s
'    s = File.ReadAllLines(Path.Combine(App.Path, "Dummy Test Files\TestWriteAllLines.txt"))
'    Assert.AreEqual 7, cArray.GetLength(s)
'    Assert.AreEqual "Four", s(3)
'    Assert.AreEqual "", s(4)
'    Assert.AreEqual "Six", s(5)
'    Assert.AreEqual "", s(6)
'End Sub

