VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoStreamTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CryptoStreamTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "WriteBlock_UsingMultiBlockWithVariousDataSizes_CallsTransformOnce"
            Test.Use 16
            Test.Use 24
            Test.Use 30
        
        Case "WriteBlock_UsingNonMultiBlockWithVariousDataSizes_CallsTransformExpectedTimes"
            Test.Use 16, 2
            Test.Use 24, 3
            Test.Use 30, 3
            
        Case "WriteBlock_WithMultipleBlockWrites_WritesToStream"
            Test.Use 16, False
            Test.Use 24, False
            Test.Use 30, False
            Test.Use 16, True
            Test.Use 24, True
            Test.Use 30, True
            
        Case "ReadBlock_UsingMultiBlockWithVariousDataSizes_CallsTransformOnce"
            Test.Use 16
            Test.Use 24
            Test.Use 30
        
        Case "ReadBlock_WithoutMultiBlockWithVariousDataSizes_CallsTransformExpectedTimes"
            Test.Use 16, 2
            Test.Use 24, 3
            Test.Use 30, 3
            
    End Select
End Sub

Public Sub Constructor_WithReadMode_SetsReadWriteMode()
    Dim Stream As CryptoStream
    Set Stream = Cor.NewCryptoStream(New MemoryStream, New FakeTransform, ReadMode)
    
    Assert.That Stream.CanRead, Iz.True
    Assert.That Stream.CanWrite, Iz.False
End Sub

Public Sub Constructor_WithWriteMode_SetsReadWriteMode()
    Dim Stream As CryptoStream
    Set Stream = Cor.NewCryptoStream(New MemoryStream, New FakeTransform, WriteMode)
    
    Assert.That Stream.CanRead, Iz.False
    Assert.That Stream.CanWrite, Iz.True
End Sub

Public Sub Constructor_WithNullStream_ThrowsArgumentNullException()
    On Error Resume Next
    
    Cor.NewCryptoStream Nothing, New FakeTransform, WriteMode
    
    AssertArgumentNullException Err, "Stream"
End Sub

Public Sub Constructor_WithNullTransform_ThrowsArgumentNullException()
    On Error Resume Next
    
    Cor.NewCryptoStream New FakeStream, Nothing, WriteMode
    
    AssertArgumentNullException Err, "Transform"
End Sub

Public Sub Constructor_WithWriteModeOnReadOnlyStream_ThrowsArgumentException()
    Dim Output As New FakeStream
    Output.CanWrite = False
    
    On Error Resume Next
    Cor.NewCryptoStream Output, New FakeTransform, WriteMode
    
    AssertArgumentException Err, "Stream"
End Sub

Public Sub Constructor_WithReadModeOnWriteOnlyStream_ThrowsArgumentException()
    Dim Output As New FakeStream
    Output.CanRead = False
    
    On Error Resume Next
    Cor.NewCryptoStream Output, New FakeTransform, ReadMode
    
    AssertArgumentException Err, "Stream"
End Sub

Public Sub WriteBlock_UsingMultiBlockWithVariousDataSizes_CallsTransformOnce(ByVal DataSize As Long)
    Dim Output As New MemoryStream
    Dim Transform As New FakeTransform
    Dim Data() As Byte
    Data = GenerateBytes(DataSize)
    Transform.CanTransformMultipleBlocks = True
    
    With Cor.NewCryptoStream(Output, Transform, WriteMode)
        .WriteBlock Data, 0, DataSize
    End With
    
    Assert.That Transform.TransformBlockCount, Iz.EqualTo(1)
    Assert.That Output.ToArray, Iz.EqualTo(Data)
End Sub

Public Sub WriteBlock_UsingNonMultiBlockWithVariousDataSizes_CallsTransformExpectedTimes(ByVal DataSize As Long, ByVal ExpectedTransformCalls As Long)
    Dim Output As New MemoryStream
    Dim Transform As New FakeTransform
    Dim Data() As Byte
    Data = GenerateBytes(DataSize)
    Transform.CanTransformMultipleBlocks = False
    
    With Cor.NewCryptoStream(Output, Transform, WriteMode)
        .WriteBlock Data, 0, DataSize
    End With
    
    Assert.That Transform.TransformBlockCount, Iz.EqualTo(ExpectedTransformCalls)
    Assert.That Output.ToArray, Iz.EqualTo(Data)
End Sub

Public Sub WriteBlock_WithMultipleBlockWrites_WritesToStream(ByVal DataSize As Long, ByVal SupportsMultiBlock As Boolean)
    Dim Output As New MemoryStream
    Dim Transform As New FakeTransform
    Dim Data() As Byte
    Data = GenerateBytes(DataSize)
    Transform.CanTransformMultipleBlocks = SupportsMultiBlock
    
    Dim Half As Long
    Half = DataSize / 2
    With Cor.NewCryptoStream(Output, Transform, WriteMode)
        .WriteBlock Data, 0, Half
        .WriteBlock Data, Half, DataSize - Half
    End With
    
    Assert.That Output.ToArray, Iz.EqualTo(Data)
End Sub

Public Sub WriteByte_WithPartiallyFilledBlock_FlushesToStream()
    Dim Output As New MemoryStream
    Dim Transform As New FakeTransform
    Dim Data() As Byte
    Data = GenerateBytes(8)
    
    With Cor.NewCryptoStream(Output, Transform, WriteMode)
        .WriteBlock Data, 0, 7
        Assert.That Transform.TransformBlockCount, Iz.EqualTo(0)
        Assert.That Output.ToArray, Iz.Empty
        
        .WriteByte Data(7)
        Assert.That Transform.TransformBlockCount, Iz.EqualTo(1)
        Assert.That Output.ToArray, Iz.EqualTo(Data)
    End With
End Sub

Public Sub FlushFinalBlock_WithPartialBlockWritten_FlushesToStream()
    Dim Output As New MemoryStream
    Dim Transform As New FakeTransform
    Dim Data() As Byte
    Data = GenerateBytes(4)
    
    With Cor.NewCryptoStream(Output, Transform, WriteMode)
        .WriteBlock Data, 0, 4
        Assert.That Transform.TransformFinalBlockCount, Iz.EqualTo(0)
        Assert.That Output.ToArray, Iz.Empty
        
        .FlushFinalBlock
        Assert.That Transform.TransformBlockCount, Iz.EqualTo(0)
        Assert.That Transform.TransformFinalBlockCount, Iz.EqualTo(1)
        Assert.That Output.ToArray, Iz.EqualTo(Data)
    End With
End Sub

Public Sub FlushFinalBlock_WhenCalledTwice_ThrowsNotSupportedException()
    Dim Output As New MemoryStream
    
    On Error Resume Next
    With Cor.NewCryptoStream(Output, New FakeTransform, WriteMode)
        .FlushFinalBlock
        .FlushFinalBlock
    End With
    
    AssertNotSupportedException Err
End Sub

Public Sub FlushFinalBlock_WhenInReadMode_ThrowsNotSupportedException()
    Dim Output As New MemoryStream
    
    On Error Resume Next
    With Cor.NewCryptoStream(Output, New FakeTransform, ReadMode)
        .FlushFinalBlock
    End With
    
    AssertNotSupportedException Err
End Sub

Public Sub FlushFinalBlock_WithCryptoStreamOutput_CallsFlushFinalBlockOnOutput()
    Dim Output As New CryptoStreamSpy
    Output.Mode = WriteMode
    
    With Cor.NewCryptoStream(Output, New FakeTransform, WriteMode)
        .FlushFinalBlock
    End With
    
    Assert.That Output.FlushFinalBlockCount, Iz.EqualTo(1)
End Sub

Public Sub ReadBlock_UsingMultiBlockWithVariousDataSizes_CallsTransformOnce(ByVal DataSize As Long)
    Dim Stream As MemoryStream
    Dim Transform As New FakeTransform
    Dim Data() As Byte
    Data = GenerateBytes(DataSize)
    Set Stream = Cor.NewMemoryStream(Data, Writable:=False)
    Transform.CanTransformMultipleBlocks = True
    
    Dim Actual As Long
    Dim ActualData() As Byte
    ReDim ActualData(0 To DataSize - 1)
    With Cor.NewCryptoStream(Stream, Transform, ReadMode)
        Actual = .ReadBlock(ActualData, 0, DataSize)
    End With
    
    Assert.That Actual, Iz.EqualTo(DataSize)
    Assert.That Transform.TransformBlockCount, Iz.EqualTo(1)
    Assert.That ActualData, Iz.EqualTo(Data)
End Sub

Public Sub ReadBlock_WithoutMultiBlockWithVariousDataSizes_CallsTransformExpectedTimes(ByVal DataSize As Long, ByVal Expected As Long)
    Dim Stream As MemoryStream
    Dim Transform As New FakeTransform
    Dim Data() As Byte
    Data = GenerateBytes(DataSize)
    Set Stream = Cor.NewMemoryStream(Data, Writable:=False)
    Transform.CanTransformMultipleBlocks = False
    
    Dim Actual As Long
    Dim ActualData() As Byte
    ReDim ActualData(0 To DataSize - 1)
    With Cor.NewCryptoStream(Stream, Transform, ReadMode)
        Actual = .ReadBlock(ActualData, 0, DataSize)
    End With
    
    Assert.That Actual, Iz.EqualTo(DataSize)
    Assert.That Transform.TransformBlockCount, Iz.EqualTo(Expected)
    Assert.That ActualData, Iz.EqualTo(Data)
End Sub

Public Sub ReadByte_WhenCalled_ReturnsNextByte()
    Dim Stream As MemoryStream
    Set Stream = Cor.NewMemoryStream(Cor.NewBytes(1, 2, 3, 4, 5), Mode:=AcquireMode)
    
    Dim Actual As Long
    With Cor.NewCryptoStream(Stream, New FakeTransform, ReadMode)
        Actual = .ReadByte
    End With
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub ReadByte_WhenPassedEndOfStream_ReturnsMinusOne()
    Dim Stream As MemoryStream
    Set Stream = Cor.NewMemoryStream(GenerateBytes(30), Mode:=AcquireMode)
    
    Dim Actual As Long
    With Cor.NewCryptoStream(Stream, New FakeTransform, ReadMode)
        Dim Data(0 To 29) As Byte
        .ReadBlock Data, 0, 30
        Actual = .ReadByte
    End With
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub
