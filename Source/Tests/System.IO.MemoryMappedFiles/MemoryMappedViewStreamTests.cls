VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryMappedViewStreamTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MemoryMappedViewStreamTests
'
Option Explicit
Implements ICategorizable

Private mIOHelper As IOHelper

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO.MemoryMappedFiles"
End Sub

Public Sub Setup()
    Set mIOHelper = New IOHelper
End Sub

Public Sub Teardown()
    Set mIOHelper = Nothing
End Sub

Public Sub Length_WithSpecifiedSize_ReturnsLengthStreamSize()
    Dim File As MemoryMappedFile
    Dim Stream As Stream
    Dim Actual As Long
    
    Set File = MakeMappedFileWithData
    Set Stream = File.CreateViewStream(Size:=75)
    
    Actual = Stream.Length
    
    Assert.That Actual, Iz.EqualTo(75)
End Sub

Public Sub ReadByte_FromSecondPosition_ReturnsOne()
    Dim File As MemoryMappedFile
    Dim Stream As Stream
    Dim Actual As Byte
    
    Set File = MakeMappedFileWithData
    Set Stream = File.CreateViewStream
    Stream.Position = 1
    
    Actual = Stream.ReadByte
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub ReadByte_WithStreamOffsetOfOne_ReturnsOne()
    Dim File As MemoryMappedFile
    Dim Stream As Stream
    Dim Actual As Byte
    
    Set File = MakeMappedFileWithData
    Set Stream = File.CreateViewStream(1)
    
    Actual = Stream.ReadByte
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub WriteByte_WhenCalled_WritesByteToStream()
    Dim File As MemoryMappedFile
    Dim Stream As Stream
    Dim Actual() As Byte
    
    Set File = MakeMappedFileWithData
    Set Stream = File.CreateViewStream
            
    Stream.WriteByte 200
    
    Set Stream = Nothing
    Set File = Nothing
    Actual = ReadBytes(0, 1)
    Assert.That Actual, Iz.EqualTo(Array(200))
End Sub

Public Sub ReadBlock_WhenCalled_ReadsData()
    Dim File As MemoryMappedFile
    Dim Stream As Stream
    Dim Actual(0 To 4) As Byte
    
    Set File = MakeMappedFileWithData
    Set Stream = File.CreateViewStream
    
    Assert.That Stream.ReadBlock(Actual, 0, 5), Iz.EqualTo(5)
    
    Assert.That Actual, Iz.EqualTo(Array(0, 1, 2, 3, 4))
End Sub

Public Sub WriteBlock_WhenCalled_WritesData()
    Dim File As MemoryMappedFile
    Dim Stream As Stream
    Dim Actual() As Byte
    Dim Data() As Byte
    
    Set File = MakeMappedFileWithData
    Set Stream = File.CreateViewStream
    Data = Cor.NewBytes(100, 99, 98, 97, 96)
    
    Stream.WriteBlock Data, 0, 5
    
    Set Stream = Nothing
    Set File = Nothing
    Actual = ReadBytes(0, 5)
    Assert.That Actual, Iz.EqualTo(Data)
End Sub

Public Sub CanSeek_WhenStreamIsClosed_ReturnsFalse()
    Dim File As MemoryMappedFile
    Dim Stream As Stream
    Dim Actual As Boolean
    
    Set File = MakeMappedFileWithData
    Set Stream = File.CreateViewStream
    
    Stream.CloseStream
    Actual = Stream.CanSeek
    
    Assert.That Actual, Iz.False
End Sub

Private Function MakeMappedFileWithData() As MemoryMappedFile
    Set MakeMappedFileWithData = MakeMappedFile(MakeData)
End Function

Private Function MakeMappedFile(ByRef Data() As Byte, Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedFile
    mIOHelper.CreateBinaryFile "File.txt", Data
    Set MakeMappedFile = MemoryMappedFile.CreateFromFile(mIOHelper.TempFileName, Access:=Access)
End Function

Private Function ReadBytes(ByVal Offset As Long, ByVal Count As Long) As Byte()
    With Cor.NewFileStream(mIOHelper.TempFileName, OpenExisting)
        Dim Data() As Byte
        ReDim Data(0 To Count - 1)
        .Position = Offset
        .ReadBlock Data, 0, Count
        ReadBytes = Data
    End With
End Function

Private Function MakeData() As Byte()
    Dim Data(0 To 99) As Byte
    Dim i As Long
    
    For i = LBound(Data) To UBound(Data)
        Data(i) = i Mod 256
    Next

    MakeData = Data
End Function
