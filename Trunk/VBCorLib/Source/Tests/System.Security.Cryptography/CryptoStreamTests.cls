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




