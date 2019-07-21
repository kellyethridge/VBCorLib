VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryMappedFileTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MemoryMappedFileTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private mIO As IOHelper

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO.MemoryMappedFiles"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "CreateFromFile_WithInvalidAccessValue_ThrowsArgumentOutOfRangeException"
            Test.Use -1
            Test.Use 6
    End Select
End Sub

Public Sub Setup()
    Set mIO = New IOHelper
End Sub

Public Sub Teardown()
    Set mIO = Nothing
End Sub

Public Sub CreateFromFile_WithEmptyMapName_ThrowsArgumentException()
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFile "c:\abc", MapName:=""
    
    AssertArgumentException Err
End Sub

Public Sub CreateFromFile_WithNegativeCapacity_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFile "c:\abc", Capacity:=-1
    
    AssertArgumentOutOfRangeException Err, "Capacity"
End Sub

Public Sub CreateFromFile_WithInvalidAccessValue_ThrowsArgumentOutOfRangeException(ByVal InvalidValue As MemoryMappedFileAccess)
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFile "c:\abc", Access:=InvalidValue
    
    AssertArgumentOutOfRangeException Err, "Access"
End Sub

Public Sub CreateFromFile_WithFileModeAppend_ThrowsArgumentException()
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFile "c:\abc", FileMode.Append
    
    AssertArgumentException Err
End Sub

Public Sub CreateFromFile_WithMemoryMappedFileAccessWriteOnly_ThrowsArgumentException()
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFile "c:\abc", Access:=WriteOnly
    
    AssertArgumentException Err
End Sub

Public Sub CreateFromFile_WithDefaultCapacityAndZeroLengthFile_ThrowsArgumentException()
    mIO.CreateFile
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFile mIO.TempFileName
    
    AssertArgumentException Err
End Sub

Public Sub CreateFromFile_WithCapacitySmallerThanFileSize_ThrowsArgumentOutOfRangeException()
    mIO.CreateFile Data:=Space$(1000)
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFile mIO.TempFileName, Capacity:=999
    
    AssertArgumentOutOfRangeException Err, "Capacity"
End Sub

Public Sub CreateFromFile_WithAccessReadOnlyAndCapacityLargerThanFile_ThrowsArgumentException()
    mIO.CreateBinaryFile Data:=CorArray.CreateInstance(vbByte, 100)
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFile mIO.TempFileName, Capacity:=200, Access:=MemoryMappedFileAccess.ReadOnly
    
    AssertArgumentException Err
End Sub

Public Sub CreateFromFile_WithCapacityLargerThanFile_IncreasesFileSizeToMatchCapacity()
    Dim FileName As String
    FileName = mIO.CreateFile(Data:="Hello")
    
    MemoryMappedFile.CreateFromFile FileName, Capacity:=100
    
    Assert.That FileLen(FileName), Iz.EqualTo(100)
End Sub

Public Sub CreateViewAccessor_WithNegativeOffset_ThrowsArgumentOutOfRangeException()
    Dim File As MemoryMappedFile
    Set File = MakeMemoryMappedFile
    
    On Error Resume Next
    File.CreateViewAccessor -1
    
    AssertArgumentOutOfRangeException Err, "Offset"
End Sub

Public Sub CreateViewAccessor_WithNegativeSize_ThrowsArgumentOutOfRangeException()
    Dim File As MemoryMappedFile
    Set File = MakeMemoryMappedFile
    
    On Error Resume Next
    File.CreateViewAccessor Size:=-1
    
    AssertArgumentOutOfRangeException Err, "Size"
End Sub

Private Function MakeMemoryMappedFile() As MemoryMappedFile
    mIO.CreateFile Data:=Space$(1000)
    Set MakeMemoryMappedFile = MemoryMappedFile.CreateFromFile(mIO.TempFileName)
End Function


