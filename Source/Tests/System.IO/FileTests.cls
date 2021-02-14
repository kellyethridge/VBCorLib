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
        Case "Exists_WithFileName_ReturnsExpected"
            Test.Use "NonExistingFile.txt", False
            Test.Use "File.txt", True
            Test.Use "AFolder", False
            Test.Use "AFolder\NonExistingFile.txt", False
            Test.Use "AFolder\SubFile.txt", True
            Test.Use "", False
            
        Case "Exists_WithParentNavigation_ReturnsExpected"
            Test.Use "File.txt", False
            Test.Use "..\File.txt", True
        
        Case "Delete_WithInvalidPath_ThrowsArgumentException"
            Test.Use ""
            Test.Use "       "
            Test.Use "c:\<\File.txt"
            
        Case "Copy_WithInvalidFileName_ThrowsArgumentException"
            Test.Use "c:\file.txt", "", "DestFileName"
            Test.Use "", "c:\file.txt", "SourceFileName"
            Test.Use "    ", "c:\file.txt", ""
            Test.Use "c:\file.txt", "    ", ""
            Test.Use "c:\>\file.txt", "c:\file2.txt", ""
            Test.Use "c:\file.txt", "c:\<file.txt", ""
            
        Case "ReadAllLines_WithLinesContainingDifferentLineSeparators_ReturnsArrayOfLines"
            Test.Use "1" & vbCrLf & "2" & vbCrLf & "3"
            Test.Use "1" & vbCr & "2" & vbCr & "3"
            Test.Use "1" & vbLf & "2" & vbLf & "3"
            
        Case "ReadAllLines_WithInvalidPath_ThrowsArgumentException"
            Test.Use ""
            Test.Use """"
            Test.Use "  "
            
        Case "WriteAllLines_WithLines_WritesToFile"
            Test.Use NewStrings()
            Test.Use NewStrings("a", "1", "2", "b")
            Test.Use NewArrayList(Array("1", "2", "3", "a", "b", "c"))
            Test.Use NewCollection("1", "2")
            
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
    Actual = ReadFile(NewFileName)
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

Public Sub Copy_WithSourceAndDestinationAsSameFile_DoesNothing()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abc")
    
    File.Copy FileName, FileName
    
    Assert.That ReadFile(FileName), Iz.EqualTo("abc")
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
    
    Assert.That FileExists(FileName), Iz.False
End Sub

Public Sub Delete_WithInvalidPath_ThrowsArgumentException(ByVal InvalidPath As String)
    On Error Resume Next
    
    File.Delete InvalidPath
    
    AssertArgumentException Err
End Sub

Public Sub ReadAllLines_WithInvalidPath_ThrowsArgumentException(ByVal Path As String)
    On Error Resume Next
    
    File.ReadAllLines Path
    
    AssertArgumentException Err
End Sub

Public Sub ReadAllLines_WithEmptyFile_ReturnsEmptyArray()
    Dim FileName As String
    FileName = mBuilder.CreateFile()
    
    Dim Actual() As String
    Actual = File.ReadAllLines(FileName)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub ReadAllLines_WithLinesContainingDifferentLineSeparators_ReturnsArrayOfLines(ByVal Lines As String)
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:=Lines)
    
    Dim Actual() As String
    Actual = File.ReadAllLines(FileName)
    
    Assert.That Actual, Iz.EqualTo(Array("1", "2", "3"))
End Sub

Public Sub ReadLines_WithEmptyFile_ReturnsNoContent()
    Dim FileName As String
    Dim Line As Variant
    Dim HasLines As Boolean
    
    FileName = mBuilder.CreateFile()
    
    For Each Line In File.ReadLines(FileName)
        HasLines = True
        Exit For
    Next
    
    Assert.That HasLines, Iz.False
End Sub

Public Sub ReadLines_WithFileContainingLines_EnumeratesLines()
    Dim FileName As String
    Dim Line As Variant
    Dim Lines As Variant
    Dim Actual As New Collection
    
    FileName = mBuilder.TempFileName
    Lines = Array("L1", "L2", "L3")
    File.WriteAllLines FileName, Lines
    
    For Each Line In File.ReadLines(FileName)
        Actual.Add Line
    Next
    
    Assert.That Actual, Iz.EqualTo(Lines)
End Sub

Public Sub ReadAllBytes_WithEmptyFile_ReturnsEmptyArray()
    Dim FileName As String
    FileName = mBuilder.CreateFile()
    
    Dim Actual() As Byte
    Actual = File.ReadAllBytes(FileName)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub ReadAllBytes_WithDataInFile_ReturnsBytes()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="1" & vbCrLf & "2" & vbCrLf & "3")
    
    Dim Actual() As Byte
    Actual = File.ReadAllBytes(FileName)
    
    Assert.That Actual, Iz.EqualTo(NewBytes(49, 13, 10, 50, 13, 10, 51))
End Sub

Public Sub ReadAllBytes_WithEmptyPath_ThrowsArgumentException()
    On Error Resume Next
    
    File.ReadAllBytes ""
    
    AssertArgumentException Err
End Sub

Public Sub ReadAllText_WithEmptyFile_ReturnsEmptyString()
    Dim FileName As String
    FileName = mBuilder.CreateFile()
    
    Dim Actual As String
    Actual = File.ReadAllText(FileName)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub ReadAllText_WithDataInFile_ReturnsDataInString()
    Const Data      As String = "1" & vbCrLf & "2" & vbCrLf & "3"
    Dim FileName    As String
    FileName = mBuilder.CreateFile(Data:=Data)
    
    Dim Actual As String
    Actual = File.ReadAllText(FileName)
    
    Assert.That Actual, Iz.EqualTo(Data)
End Sub

Public Sub WriteAllBytes_WithEmptyData_CreatesEmptyFile()
    Dim FileName As String
    FileName = mBuilder.CreateFile()
    
    File.WriteAllbytes FileName, NewBytes()
    
    Dim Actual() As Byte
    Actual = File.ReadAllBytes(FileName)
    Assert.That Actual, Iz.Empty
End Sub

Public Sub WriteAllBytes_WithData_WritesToFile()
    Dim FileName    As String
    Dim Data()      As Byte
    Data = NewBytes(1, 2, 3, 4, 5)
    FileName = mBuilder.CreateFile(Data:="Overwrite me")
    
    File.WriteAllbytes FileName, Data
    
    Dim Actual() As Byte
    Actual = File.ReadAllBytes(FileName)
    Assert.That Actual, Iz.EqualTo(Data)
End Sub

Public Sub WriteAllLines_WithLines_WritesToFile(ByVal Contents As Variant)
    Dim FileName As String
    FileName = mBuilder.CreateFile()
    
    File.WriteAllLines FileName, Contents
    
    Dim Actual() As String
    Actual = File.ReadAllLines(FileName)
    Assert.That Actual, Iz.EqualTo(Contents).AsCollection
End Sub

Public Sub WriteAllText_WithExistingFile_OverwritesFile()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abc")
    
    File.WriteAllText FileName, "Hello"
    
    Dim Actual As String
    Actual = File.ReadAllText(FileName)
    Assert.That Actual, Iz.EqualTo("Hello")
End Sub

Public Sub WriteAllText_WithNonExistingFile_CreatesNewFile()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "MyFile.txt")
    
    File.WriteAllText FileName, "Hello world"
    
    Dim Actual As String
    Actual = File.ReadAllText(FileName)
    Assert.That Actual, Iz.EqualTo("Hello world")
End Sub

Public Sub Attributes_WhenSet_AreSetToFile()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abc")
    
    File.SetAttributes FileName, FileAttributes.Hidden
    
    Dim Actual As FileAttributes
    Actual = File.GetAttributes(FileName)
    Assert.That Actual, Iz.EqualTo(FileAttributes.Hidden)
End Sub

Public Sub Move_WithNonExistingDestination_MovesSourceToDest()
    Dim SourceFileName As String
    Dim DestinationFileName As String
    SourceFileName = mBuilder.CreateFile(Data:="abc")
    DestinationFileName = Path.Combine(mBuilder.TempFolder, "File2.txt")
    
    File.Move SourceFileName, DestinationFileName
    
    Dim Actual As String
    Actual = File.ReadAllText(DestinationFileName)
    Assert.That Actual, Iz.EqualTo("abc")
    Assert.That File.Exists(SourceFileName), Iz.False
End Sub

Public Sub Move_WithSameSourceAndDestination_DoesNothing()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abc")
    
    File.Move FileName, FileName
    
    Assert.That File.Exists(FileName), Iz.True
End Sub

Public Sub Move_WithExistingDestination_ThrowsIOException()
    Dim SourceFileName As String
    Dim DestinationFileName As String
    SourceFileName = mBuilder.CreateFile(Data:="abc")
    DestinationFileName = mBuilder.CreateFile(FileName:="File2.txt", Data:="def")
    
    On Error Resume Next
    File.Move SourceFileName, DestinationFileName
    
    AssertIOException Err
End Sub

Public Sub AppendAllText_WithNonExistingFile_CreatesNewFile()
    Const Data As String = "ABC"
    
    File.AppendAllText mBuilder.TempFileName, Data
    
    Dim Actual As String
    Actual = File.ReadAllText(mBuilder.TempFileName)
    Assert.That Actual, Iz.EqualTo(Data)
End Sub

Public Sub AppendAllText_WithExistingFile_AppendToEndOfFile()
    mBuilder.CreateFile Data:="abc"
    
    File.AppendAllText mBuilder.TempFileName, "def"
    
    Dim Actual As String
    Actual = File.ReadAllText(mBuilder.TempFileName)
    Assert.That Actual, Iz.EqualTo("abcdef")
End Sub

Public Sub AppendAllLines_WithNonExistingFile_CreatesNewFile()
    Dim Lines As Variant
    
    Lines = Array("Line2", "Line3")
    
    File.AppendAllLines mBuilder.TempFileName, Lines
    
    Dim Actual() As String
    Actual = File.ReadAllLines(mBuilder.TempFileName)
    Assert.That Actual, Iz.EqualTo(Array("Line2", "Line3"))
End Sub

Public Sub AppendAllLines_WithExistingFile_AppendsToFile()
    Dim Lines As Variant
    
    Lines = Array("Line2", "Line3")
    mBuilder.CreateFile Data:="Line1" & vbCrLf
    
    File.AppendAllLines mBuilder.TempFileName, Lines
    
    Dim Actual() As String
    Actual = File.ReadAllLines(mBuilder.TempFileName)
    Assert.That Actual, Iz.EqualTo(Array("Line1", "Line2", "Line3"))
End Sub

Public Sub OpenFile_WithModeAppendAndDefaultAccess_ReturnsFileStreamWithWriteAccessOnly()
    Dim Actual As FileStream
    
    mBuilder.CreateFile Data:="abc"
    
    Set Actual = File.OpenFile(mBuilder.TempFileName, FileMode.Append)
    
    Assert.That Actual.CanWrite, Iz.True
    Assert.That Actual.CanRead, Iz.False
End Sub

Public Sub OpenFile_WithModeOpenExistingAndDefaultAccess_ReturnsFileStreamWithReadWriteAccess()
    Dim Actual As FileStream
    
    mBuilder.CreateFile Data:="abc"
    
    Set Actual = File.OpenFile(mBuilder.TempFileName, FileMode.OpenExisting)
    
    Assert.That Actual.CanWrite, Iz.True
    Assert.That Actual.CanRead, Iz.True
End Sub




















