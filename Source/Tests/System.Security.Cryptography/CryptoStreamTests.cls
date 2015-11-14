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
        Case "WriteBlock_WithTransformType_WritesExpectedTransforms"
            Test.Use False, 2
            Test.Use True, 1
            
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

Public Sub WriteBlock_WithTransformType_WritesExpectedTransforms(ByVal SupportsMultipleBlock As Boolean, ByVal ExpectedTransformCount As Long)
    Dim Output As New MemoryStream
    Dim Transform As New FakeTransform
    Dim Data() As Byte
    Data = GenerateBytes(16)
    Transform.CanTransformMultipleBlocks = SupportsMultipleBlock
    
    With Cor.NewCryptoStream(Output, Transform, WriteMode)
        .WriteBlock Data, 0, 16
    End With
    
    Assert.That Transform.TransformBlockCount, Iz.EqualTo(ExpectedTransformCount)
    Assert.That Output.ToArray, Iz.EqualTo(Data)
End Sub


