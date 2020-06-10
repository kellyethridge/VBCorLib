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
            Test.Use LessThanMemoryMappedFileAccess
            Test.Use MoreThanMemoryMappedFileAccess
            
        Case "CreateFromFileStream_WithOutOfRangeAccessValue_ThrowsArgumentOutOfRangeException"
            Test.Use LessThanMemoryMappedFileAccess
            Test.Use MoreThanMemoryMappedFileAccess
        
        Case "CreateNew_WithAccessOutOfRange_ThrowsArgumentOutOfRangeException"
            Test.Use LessThanMemoryMappedFileAccess
            Test.Use MoreThanMemoryMappedFileAccess
        
        Case "CreateOrOpen_WithAccessOutOfRange_ThrowsArgumentOutOfRangeException"
            Test.Use LessThanMemoryMappedFileAccess
            Test.Use MoreThanMemoryMappedFileAccess
        
        Case "OpenExisting_WithDesiredAccessRightsOutOfRange_ThrowsArgumentOutOfRangeException"
            Test.Use &H10
            Test.Use &H100000
        
        Case "CreateViewStream_WithAccessOutOfRange_ThrowsArgumentOutOfRangeException"
            Test.Use LessThanMemoryMappedFileAccess
            Test.Use MoreThanMemoryMappedFileAccess
        
        Case "CreateViewStream_WithAccess_SetsReadAndWrite"
            Test.Use MemoryMappedFileAccess.CopyOnWrite, True, True
            Test.Use MemoryMappedFileAccess.ReadExecute, True, False
            Test.Use MemoryMappedFileAccess.ReadOnly, True, False
            Test.Use MemoryMappedFileAccess.ReadWrite, True, True
            Test.Use MemoryMappedFileAccess.ReadWriteExecute, True, True
            Test.Use MemoryMappedFileAccess.WriteOnly, False, True
            
    End Select
End Sub

Private Property Get LessThanMemoryMappedFileAccess() As Long
    LessThanMemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite - 1
End Property

Private Property Get MoreThanMemoryMappedFileAccess() As Long
    MoreThanMemoryMappedFileAccess = MemoryMappedFileAccess.ReadWriteExecute + 1
End Property

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
    
    AssertArgumentException Err, "Mode"
End Sub

Public Sub CreateFromFile_WithMemoryMappedFileAccessWriteOnly_ThrowsArgumentException()
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFile "c:\abc", Access:=WriteOnly
    
    AssertArgumentException Err, "Access"
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

Public Sub CreateViewAccessor_WhenCalled_ReturnsAccessor()
    Dim File As MemoryMappedFile
    Dim Accessor As MemoryMappedViewAccessor
    
    Set File = MakeMemoryMappedFile
    Set Accessor = File.CreateViewAccessor
    
    Assert.That Accessor, Iz.Not.Nothing
End Sub

Public Sub CreateFromFileStream_WithFileStreamAsNothing_ThrowsArgumentNullException()
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFileStream Nothing, vbNullString, 0, MemoryMappedFileAccess.ReadWrite, True
    
    AssertArgumentNullException Err, "FileStream"
End Sub

Public Sub CreateFromFileStream_WithEmptyMapName_ThrowsArgumentException()
    On Error Resume Next

    MemoryMappedFile.CreateFromFileStream MakeFileStream, "", 0, MemoryMappedFileAccess.ReadWrite, True

    AssertArgumentException Err
End Sub

Public Sub CreateFromFileStream_WithNegativeCapacity_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFileStream MakeFileStream, vbNullString, -1, MemoryMappedFileAccess.ReadWrite, True
    
    AssertArgumentOutOfRangeException Err, "Capacity"
End Sub

Public Sub CreateFromFileStream_WithOutOfRangeAccessValue_ThrowsArgumentOutOfRangeException(ByVal InvalidValue As MemoryMappedFileAccess)
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFileStream MakeFileStream, vbNullString, 0, InvalidValue, True
    
    AssertArgumentOutOfRangeException Err, "Access"
End Sub

Public Sub CreateFromFileStream_WithDefaultCapacityAndZeroLengthFile_ThrowsArgumentException()
    Dim Stream As FileStream
    mIO.CreateFile
    Set Stream = Cor.NewFileStream(mIO.TempFileName, OpenExisting)
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFileStream Stream, vbNullString, 0, MemoryMappedFileAccess.ReadWrite, True
        
    AssertArgumentException Err
End Sub

Public Sub CreateFromFileStream_WithWriteOnlyAccess_ThrowsArgumentException()
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFileStream MakeFileStream, vbNullString, 0, MemoryMappedFileAccess.WriteOnly, True
    
    AssertArgumentException Err, "Access"
End Sub

Public Sub CreateFromFileStream_WithReadOnlyAccessAndCapacityGreaterThanStreamLength_ThrowsArgumentException()
    Dim Stream As FileStream
    Set Stream = MakeFileStream
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFileStream Stream, vbNullString, Stream.Length + 1, MemoryMappedFileAccess.ReadOnly, True
    
    AssertArgumentException Err
End Sub

Public Sub CreateFromFileStream_WithCapacityLessThanStreamLength_ThrowsArgumentException()
    Dim Stream As FileStream
    Set Stream = MakeFileStream
    On Error Resume Next
    
    MemoryMappedFile.CreateFromFileStream Stream, vbNullString, Stream.Length - 1, MemoryMappedFileAccess.ReadWrite, True
    
    AssertArgumentOutOfRangeException Err, "Capacity"
End Sub

Public Sub CreateNew_WithEmptyMapName_ThrowsArgumentException()
    On Error Resume Next
    
    MemoryMappedFile.CreateNew "", 0
    
    AssertArgumentException Err
End Sub

Public Sub CreateNew_WithZeroCapacity_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    MemoryMappedFile.CreateNew "a", 0
    
    AssertArgumentOutOfRangeException Err, "Capacity"
End Sub

Public Sub CreateNew_WithNegativeCapacity_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    MemoryMappedFile.CreateNew "a", -1
    
    AssertArgumentOutOfRangeException Err, "Capacity"
End Sub

Public Sub CreateNew_WithAccessOutOfRange_ThrowsArgumentOutOfRangeException(ByVal Access As MemoryMappedFileAccess)
    On Error Resume Next
    
    MemoryMappedFile.CreateNew "a", 1, Access
    
    AssertArgumentOutOfRangeException Err, "Access"
End Sub

Public Sub CreateNew_WithWriteOnlyAccess_ThrowsArgumentException()
    On Error Resume Next
    
    MemoryMappedFile.CreateNew "a", 1, MemoryMappedFileAccess.WriteOnly
    
    AssertArgumentException Err, "Access"
End Sub

Public Sub CreateNew_WhenCalled_ReturnsMemoryMappedFile()
    Dim File As MemoryMappedFile
    
    Set File = MemoryMappedFile.CreateNew("abc", 1)
        
    Assert.That File, Iz.Not.Nothing
    Assert.That File.SafeMemoryMappedFileHandle.IsInvalid, Iz.False
End Sub

Public Sub OpenExising_WithNullMapName_ThrowsArgumentNullException()
    On Error Resume Next
    
    MemoryMappedFile.OpenExisting vbNullString
    
    AssertArgumentNullException Err, "MapName"
End Sub

Public Sub OpenExisting_WithEmptyMapName_ThrowsArgumentException()
    On Error Resume Next
    
    MemoryMappedFile.OpenExisting ""
    
    AssertArgumentException Err
End Sub

Public Sub OpenExisting_WithDesiredAccessRightsOutOfRange_ThrowsArgumentOutOfRangeException(ByVal DesiredAccessRights As MemoryMappedFileRights)
    On Error Resume Next
    
    MemoryMappedFile.OpenExisting "abc", DesiredAccessRights
    
    AssertArgumentOutOfRangeException Err, "DesiredAccessRights"
End Sub


Public Sub OpenExisting_WithNonExistingMapName_ThrowsFileNotFoundException()
    On Error Resume Next

    MemoryMappedFile.OpenExisting "NotFound"

    AssertFileNotFoundException Err
End Sub

Public Sub OpenExisting_WhenCalled_ReturnsMemoryMappedFile()
    Dim Existing As MemoryMappedFile
    Dim File As MemoryMappedFile

    Set Existing = MemoryMappedFile.CreateNew("TestFile", 1)
    Set File = MemoryMappedFile.OpenExisting("TestFile")

    Assert.That File, Iz.Not.Nothing
    Assert.That File.SafeMemoryMappedFileHandle.IsInvalid, Iz.False
End Sub

Public Sub CreateOrOpen_WithNullMapName_ThrowsArgumentNullException()
    On Error Resume Next
    
    MemoryMappedFile.CreateOrOpen vbNullString, 0
    
    AssertArgumentNullException Err, "MapName"
End Sub

Public Sub CreateOrOpen_WithEmptyMapName_ThrowsArgumentException()
    On Error Resume Next
    
    MemoryMappedFile.CreateOrOpen "", 0
    
    AssertArgumentException Err
End Sub

Public Sub CreateOrOpen_WithCapacityEqualToZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    MemoryMappedFile.CreateOrOpen "abc", 0
    
    AssertArgumentOutOfRangeException Err, "Capacity"
End Sub

Public Sub CreateOrOpen_WithCapacityLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    MemoryMappedFile.CreateOrOpen "abc", -1
    
    AssertArgumentOutOfRangeException Err, "Capacity"
End Sub

Public Sub CreateOrOpen_WithAccessOutOfRange_ThrowsArgumentOutOfRangeException(ByVal Access As MemoryMappedFileAccess)
    On Error Resume Next
    
    MemoryMappedFile.CreateOrOpen "abc", 1, Access
    
    AssertArgumentOutOfRangeException Err, "Access"
End Sub

Public Sub CreateOrOpen_WithOutExistingMappingObject_ReturnsValidMemoryMappedFile()
    Dim File As MemoryMappedFile
    
    Set File = MemoryMappedFile.CreateOrOpen("Created", 1)
    
    Assert.That File.SafeMemoryMappedFileHandle.IsInvalid, Iz.False
End Sub

Public Sub CreateOrOpen_WithExistingMappingObject_ReturnsValidMemoryMappedFile()
    Dim Existing As MemoryMappedFile
    Dim File As MemoryMappedFile
    Dim ExistingAccessor As MemoryMappedViewAccessor
    Dim FileAccessor As MemoryMappedViewAccessor
    
    Set Existing = MemoryMappedFile.CreateNew("Existing", 4)
    Set File = MemoryMappedFile.CreateOrOpen("Existing", 4)
    
    Set ExistingAccessor = Existing.CreateViewAccessor
    Set FileAccessor = File.CreateViewAccessor
    
    Assert.That File.SafeMemoryMappedFileHandle.IsInvalid, Iz.False
    
    FileAccessor.WriteValue 0, 56
    Assert.That ExistingAccessor.ReadInt32(0), Iz.EqualTo(56)
End Sub

Public Sub CreateViewAccessor_WhenUsingSuppliedFileStream_ReturnsAccessor()
    Dim File As MemoryMappedFile
    Dim Stream As FileStream
    Dim Accessor As MemoryMappedViewAccessor
    
    Set Stream = MakeFileStream
    Set File = MemoryMappedFile.CreateFromFileStream(Stream, vbNullString, 0, MemoryMappedFileAccess.ReadWrite, True)
    
    Set Accessor = File.CreateViewAccessor
    
    Assert.That Accessor, Iz.Not.Nothing
End Sub

Public Sub CreateViewAccessor_WithOffset_ReturnsAccessorUsingOffset()
    Dim File As MemoryMappedFile
    Dim Accessor As MemoryMappedViewAccessor
    
    Set File = MakeMemoryMappedFile
    Set Accessor = File.CreateViewAccessor(10)
    
    Assert.That Accessor.ReadInt64(0), Iz.EqualTo(NewInt64(&HD0C0B0A, &H11100F0E))
End Sub

Public Sub CreateViewStream_WithNegativeOffset_ThrowsArgumentOutOfRangeException()
    Dim File As MemoryMappedFile
    Set File = MakeMemoryMappedFile
    
    On Error Resume Next
    File.CreateViewStream -1
    
    AssertArgumentOutOfRangeException Err, "Offset"
End Sub

Public Sub CreateViewStream_WithNegativeSize_ThrowsArgumentOutOfRangeException()
    Dim File As MemoryMappedFile
    Set File = MakeMemoryMappedFile
    
    On Error Resume Next
    File.CreateViewStream 0, -1
    
    AssertArgumentOutOfRangeException Err, "Size"
End Sub

Public Sub CreateViewStream_WithAccessOutOfRange_ThrowsArgumentOutOfRangeException(ByVal Access As MemoryMappedFileAccess)
    Dim File As MemoryMappedFile
    Set File = MakeMemoryMappedFile
    On Error Resume Next
    
    File.CreateViewStream 0, 1, Access
    
    AssertArgumentOutOfRangeException Err, "Access"
End Sub

Public Sub CreateViewStream_WithAccess_SetsReadAndWrite(ByVal Access As MemoryMappedFileAccess, ByVal CanRead As Boolean, ByVal CanWrite As Boolean)
    Dim File As MemoryMappedFile
    Dim Stream As Stream
    Set File = MemoryMappedFile.CreateNew("Test", 4, MemoryMappedFileAccess.ReadWriteExecute)
    
    Set Stream = File.CreateViewStream(Access:=Access)
    
    Assert.That Stream.CanRead, Iz.EqualTo(CanRead)
    Assert.That Stream.CanWrite, Iz.EqualTo(CanWrite)
End Sub


Private Function MakeMemoryMappedFile(Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedFile
    mIO.CreateBinaryFile Data:=MakeData
    Set MakeMemoryMappedFile = MemoryMappedFile.CreateFromFile(mIO.TempFileName, Access:=Access)
End Function

Private Function MakeFileStream(Optional ByVal Access As FileAccess = FileAccess.ReadWriteAccess) As FileStream
    mIO.CreateBinaryFile Data:=MakeData
    Set MakeFileStream = Cor.NewFileStream(mIO.TempFileName, OpenExisting, Access)
End Function

Private Function MakeData() As Byte()
    Dim Data(0 To 99) As Byte
    Dim i As Long
    
    For i = LBound(Data) To UBound(Data)
        Data(i) = i Mod 256
    Next

    MakeData = Data
End Function
