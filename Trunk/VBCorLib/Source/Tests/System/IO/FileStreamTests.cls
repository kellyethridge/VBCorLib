VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileStreamTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' FileStreamTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Const OneChar   As Long = 49
Private Const TwoChar   As Long = 50
Private Const ThreeChar As Long = 51
Private Const FourChar  As Long = 52
Private Const FiveChar  As Long = 53

Private mBuilder As IOHelper

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ReadBlock_WhenCalled_ReadsData"
            Test.Use 0, 0, NewBytes()
            Test.Use 1, 1, NewBytes(49)
            Test.Use 2, 2, NewBytes(49, 50)
            Test.Use 3, 3, NewBytes(49, 50, 51)
            Test.Use 4, 4, NewBytes(49, 50, 51, 52)
            Test.Use 5, 5, NewBytes(49, 50, 51, 52, 53)
            
        Case "SeekPosition_WithValues_MovesToFilePosition"
            Test.Use 0, SeekOrigin.FromBeginning, OneChar
            Test.Use 4, SeekOrigin.FromBeginning, FiveChar
            Test.Use 2, SeekOrigin.FromBeginning, ThreeChar
            Test.Use 0, SeekOrigin.FromEnd, -1
            Test.Use -1, SeekOrigin.FromEnd, FiveChar
            Test.Use -4, SeekOrigin.FromEnd, TwoChar
            Test.Use -5, SeekOrigin.FromEnd, OneChar
            
        Case "Position_WhenSet_MovesToFilePosition"
            Test.Use 0, OneChar
            Test.Use 3, FourChar
            Test.Use 4, FiveChar
            Test.Use 5, -1
            
    End Select
End Sub

Public Sub Setup()
    Set mBuilder = New IOHelper
End Sub

Public Sub Teardown()
    Set mBuilder = Nothing
End Sub

Public Sub NewFileStream_WithEmptyPath_ThrowsArgumentException()
    On Error Resume Next
    
    NewFileStream "", Create
    
    AssertArgumentException Err, "Path"
End Sub

Public Sub CanRead_WithFileModeAppend_ReturnsFalse()
    Dim FileName As String
    FileName = mBuilder.CreateFile("Test.txt", "Hello")
    
    With NewFileStream(FileName, FileMode.Append)
        Assert.That .CanRead, Iz.False
        Assert.That .CanWrite, Iz.True
        Assert.That .CanSeek, Iz.True
    End With
End Sub

Public Sub CanRead_WithFileModeCreate_ReturnsTrue()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "Test.txt")
    
    With NewFileStream(FileName, FileMode.Create)
        Assert.That .CanRead, Iz.True
        Assert.That .CanWrite, Iz.True
        Assert.That .CanSeek, Iz.True
    End With
End Sub

Public Sub CanWrite_WithFileModeCreate_ReturnsTrue()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "Test.txt")
    
    Dim Actual As Boolean
    With NewFileStream(FileName, FileMode.Create)
        Actual = .CanWrite
    End With
    
    Assert.That Actual, Iz.True
End Sub

Public Sub NewFileStream_WithBufferSizeLessThanOne_ThrowsArgumentOutOfRangeException()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "Test.txt")
    
    On Error Resume Next
    NewFileStream FileName, Create, BufferSize:=0
    
    AssertArgumentOutOfRangeException Err, "BufferSize"
End Sub

Public Sub ReadByte_WhenCalled_ReturnsBytes()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12")
    
    With NewFileStream(FileName, FileMode.OpenExisting)
        Assert.That .ReadByte, Iz.EqualTo(Asc("1"))
        Assert.That .ReadByte, Iz.EqualTo(Asc("2"))
        Assert.That .ReadByte, Iz.EqualTo(-1)
    End With
End Sub

Public Sub NewFileStream_WithAppendFileModeAndReadFileAccess_ThrowsArgumentException()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    
    On Error Resume Next
    With NewFileStream(FileName, FileMode.Append, FileAccess.ReadAccess)
    End With
    
    WriteLine Err.Description
    AssertArgumentException Err
End Sub

Public Sub Length_WhenCalled_ReturnsFileLength()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12345678")
    
    Dim Actual As Long
    With NewFileStream(FileName, OpenExisting)
        Actual = .Length
    End With
    
    Assert.That Actual, Iz.EqualTo(8)
End Sub

Public Sub ReadBlock_WithNullBytes_ThrowsArgumentNullException()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12345")
    
    On Error Resume Next
    Dim NullBytes() As Byte
    With NewFileStream(FileName, OpenExisting)
        .ReadBlock NullBytes, 0, 1
    End With
    
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub ReadBlock_WithOffsetLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12345")
    
    On Error Resume Next
    Dim Bytes(1 To 5) As Byte
    With NewFileStream(FileName, OpenExisting)
        .ReadBlock Bytes, 0, 1
    End With

    AssertArgumentOutOfRangeException Err, "Offset"
End Sub

Public Sub ReadBlock_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12345")
    
    On Error Resume Next
    Dim Bytes(0 To 3) As Byte
    With NewFileStream(FileName, OpenExisting)
        .ReadBlock Bytes, 1, -1
    End With
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub ReadBlock_WithOffsetAndCountExtendingPassedEndOfBytes_ThrowsArgumentException()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12345")
    
    On Error Resume Next
    Dim Bytes(0 To 3) As Byte
    With NewFileStream(FileName, OpenExisting)
        .ReadBlock Bytes, 1, 4
    End With

    AssertArgumentException Err
End Sub

Public Sub ReadBlock_WhenCalled_ReadsData(ByVal CountToRead As Long, ByVal ExpectedCount As Long, ByRef ExpectedBytes() As Byte)
    Dim ActualBytes()   As Byte
    Dim ActualCount     As Long
    Dim FileName        As String
    FileName = mBuilder.CreateFile(Data:="12345")
    
    ActualBytes = CorArray.CreateInstance(vbByte, CountToRead)
    With NewFileStream(FileName, OpenExisting)
        ActualCount = .ReadBlock(ActualBytes, 0, CountToRead)
    End With
    
    Assert.That ActualCount, Iz.EqualTo(ExpectedCount)
    Assert.That ActualBytes, Iz.EqualTo(ExpectedBytes)
End Sub

Public Sub ReadBlock_WithCountPassedEnd_ReadsToEnd()
    Dim ActualBytes()   As Byte
    Dim ActualCount     As Long
    Dim FileName        As String
    FileName = mBuilder.CreateFile(Data:="12345")
    
    ReDim ActualBytes(0 To 5)
    With NewFileStream(FileName, OpenExisting)
        ActualCount = .ReadBlock(ActualBytes, 0, 6)
    End With
    
    Assert.That ActualCount, Iz.EqualTo(5)
    Assert.That ActualBytes, Iz.EqualTo(NewBytes(49, 50, 51, 52, 53, 0))
End Sub

Public Sub ReadBlock_WithLargerThanBuffer_ReadsData()
    Dim Data As String
    Data = GenerateString(2000)
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:=Data)
    
    Dim ActualBytes(0 To 1999) As Byte
    Dim ActualCount As Long
    With NewFileStream(FileName, OpenExisting, BufferSize:=512)
        ActualCount = .ReadBlock(ActualBytes, 0, 2000)
    End With
    
    Assert.That ActualCount, Iz.EqualTo(2000)
    Assert.That StrConv(ActualBytes, vbUnicode), Iz.EqualTo(Data)
End Sub

Public Sub SeekPosition_WithValues_MovesToFilePosition(ByVal Offset As Long, ByVal Origin As SeekOrigin, ByVal Expected As Long)
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12345")
    
    Dim Actual As Long
    With NewFileStream(FileName, OpenExisting, ReadAccess)
        .ReadByte   ' cause buffer to be filled to test seeking within current buffer.
        .SeekPosition Offset, Origin
        Actual = .ReadByte
    End With
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SeekPosition_ToBeforeStartOfStream_ThrowsArgumentException()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12345")
        
    On Error Resume Next
    With NewFileStream(FileName, OpenExisting)
        .SeekPosition -1, FromBeginning
    End With
    
    WriteLine Err.Description
    AssertArgumentException Err
End Sub

Public Sub Position_WhenSet_MovesToFilePosition(ByVal NewPosition As Long, ByVal Expected As Long)
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12345")
    
    Dim Actual As Long
    With NewFileStream(FileName, OpenExisting)
        .Position = NewPosition
        Actual = .ReadByte
    End With
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Position_ToBeforeStartOfStream_ThrowsArgumentException()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12345")
        
    On Error Resume Next
    With NewFileStream(FileName, OpenExisting)
        .Position = -1
    End With
    
    WriteLine Err.Description
    AssertArgumentException Err
End Sub

Public Sub Position_WhenSetBeforeStartOfAppendPoint_ThrowsIOException()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="abcdef")
    
    On Error Resume Next
    With NewFileStream(FileName, Append)
        .Position = 0
    End With
    
    AssertIOException Err
End Sub

Public Sub WriteBlock_WhenCreatingNew_WritesNewFile()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "File.txt")
    
    Dim Data() As Byte
    Data = NewBytes(1, 2, 3, 4, 5, 6)
    With NewFileStream(FileName, CreateNew)
        .WriteBlock Data, 0, 6
    End With
    
    Dim Actual() As Byte
    Actual = mBuilder.ReadBinaryFile(FileName)
    Assert.That Actual, Iz.EqualTo(Data)
End Sub

Public Sub WriteBlock_WhenAppending_AppendsToFile()
    Dim FileName As String
    FileName = mBuilder.CreateBinaryFile(Data:=NewBytes(1, 2, 3))
    
    With NewFileStream(FileName, Append)
        .WriteBlock NewBytes(4, 5, 6), 0, 3
    End With
    
    Dim Actual() As Byte
    Actual = mBuilder.ReadBinaryFile(FileName)
    Assert.That Actual, Iz.EqualTo(NewBytes(1, 2, 3, 4, 5, 6))
End Sub

Public Sub WriteBlock_WhenCreating_OverwritesFile()
    Dim FileName As String
    FileName = mBuilder.CreateBinaryFile(Data:=NewBytes(1, 2, 3))
    
    With NewFileStream(FileName, Create)
        .WriteBlock NewBytes(4, 5, 6), 0, 3
    End With
    
    Dim Actual() As Byte
    Actual = mBuilder.ReadBinaryFile(FileName)
    Assert.That Actual, Iz.EqualTo(NewBytes(4, 5, 6))
End Sub

Public Sub WriteBlock_WhenOpenOrCreatingAndFileExists_OverwritesFile()
    Dim FileName As String
    FileName = mBuilder.CreateBinaryFile(Data:=NewBytes(1, 2, 3))
    
    With NewFileStream(FileName, OpenOrCreate)
        .WriteBlock NewBytes(4, 5, 6), 0, 3
    End With
    
    Dim Actual() As Byte
    Actual = mBuilder.ReadBinaryFile(FileName)
    Assert.That Actual, Iz.EqualTo(NewBytes(4, 5, 6))
End Sub

Public Sub WriteBlock_WhenOpenOrCreateAndFileDoesNotExist_WritesNewFile()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "File.txt")
    
    With NewFileStream(FileName, OpenOrCreate)
        .WriteBlock NewBytes(4, 5, 6), 0, 3
    End With
    
    Dim Actual() As Byte
    Actual = mBuilder.ReadBinaryFile(FileName)
    Assert.That Actual, Iz.EqualTo(NewBytes(4, 5, 6))
End Sub

Public Sub WriteBlock_WithNullArray_ThrowsArgumentNullException()
    On Error Resume Next
    
    Dim NullBytes() As Byte
    With NewFileStream(mBuilder.TempFileName, Create)
        .WriteBlock NullBytes, 0, 1
    End With
    
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub WriteBlock_WithOffsetLessThanLBound_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Dim Bytes(5 To 10) As Byte
    With NewFileStream(mBuilder.TempFileName, Create)
        .WriteBlock Bytes, 4, 1
    End With
    
    AssertArgumentOutOfRangeException Err, "Offset"
End Sub

Public Sub WriteBlock_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Dim Bytes(5 To 10) As Byte
    With NewFileStream(mBuilder.TempFileName, Create)
        .WriteBlock Bytes, 5, -1
    End With
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub WriteBlock_WithOffsetAndCountNotAValidPositionInArray_ThrowsArgumentException()
    On Error Resume Next
    
    Dim Bytes(0 To 4) As Byte
    With NewFileStream(mBuilder.TempFileName, Create)
        .WriteBlock Bytes, 1, 5
    End With
    
    AssertArgumentException Err
End Sub

Public Sub WriteByte_WhenCalled_WritesToStream()
    Dim FileName As String
    FileName = mBuilder.CreateBinaryFile(Data:=NewBytes(1, 2, 3))
    
    With NewFileStream(FileName, Append)
        .WriteByte 4
    End With
    
    Dim Actual() As Byte
    Actual = mBuilder.ReadBinaryFile(FileName)
    Assert.That Actual, Iz.EqualTo(NewBytes(1, 2, 3, 4))
End Sub

Public Sub SetLength_WithNegativeValue_ThrowsArgumentOutOfException()
    On Error Resume Next
    
    With NewFileStream(mBuilder.TempFileName, Create)
        .SetLength -1
    End With
    
    AssertArgumentOutOfRangeException Err, "Value"
End Sub

Public Sub SetLength_WithValueLessThanFileSize_TruncatesFile()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="1234567890")
    
    With NewFileStream(FileName, OpenExisting)
        .SetLength 5
    End With
    
    Dim Actual As String
    Actual = mBuilder.ReadFile(FileName)
    Assert.That Actual, Iz.EqualTo("12345")
End Sub

Public Sub SetLength_WithValueGreaterThanFileSize_ExtendsFileWithNulls()
    Dim FileName As String
    FileName = mBuilder.CreateBinaryFile(Data:=NewBytes(1, 2, 3, 4))
    
    With NewFileStream(FileName, OpenExisting)
        .SetLength 9
    End With
    
    Dim Actual() As Byte
    Actual = mBuilder.ReadBinaryFile(FileName)
    Assert.That Actual, Iz.EqualTo(NewBytes(1, 2, 3, 4, 0, 0, 0, 0, 0))
End Sub

Public Sub SetLength_BeforeBeginningOfAppendStart_ThrowsIOException()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="1234567890")
    
    On Error Resume Next
    With NewFileStream(FileName, Append)
        .SetLength 9
    End With
    
    AssertIOException Err
End Sub

Public Sub NewFileStream_WithNonExistingFile_ThrowsFileNotFoundException()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "NonExisting.txt")
    
    On Error Resume Next
    With NewFileStream(FileName, OpenExisting)
    End With
    
    AssertFileNotFoundException Err, FileName
End Sub

Public Sub NewFileStream_WithCreatingNewAndFileExists_ThrowsIOException()
    Dim FileName As String
    FileName = mBuilder.CreateFile
    
    On Error Resume Next
    With NewFileStream(FileName, CreateNew)
    End With
    
    AssertIOException Err
End Sub

