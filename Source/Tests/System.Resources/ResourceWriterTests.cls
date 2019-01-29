VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceWriterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ResourceWriterTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "NewResourceWriter_WithNonSupportedSource_ThrowsArgumentException"
            Test.Use 1
            Test.Use New Collection
            
    End Select
End Sub

Public Sub NewResourceWriter_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    
    Cor.NewResourceWriter Nothing
    
    AssertArgumentNullException Err, "Source"
End Sub

Public Sub NewResourceWriter_WithNonSupportedSource_ThrowsArgumentException(ByRef Source As Variant)
    On Error Resume Next
    
    Cor.NewResourceWriter Source
    
    AssertArgumentException Err, "Source"
End Sub

Public Sub Generate_WithNoResources_WriteBlankHeaderOnly()
    Dim Writer As ResourceWriter
    Dim Output As New MemoryStream
    Set Writer = Cor.NewResourceWriter(Output)
    
    Writer.Generate
    
    Dim Actual() As Byte
    Actual = Output.ToArray
    Dim Expected() As Byte
    Expected = Cor.NewBytes(0, 0, 0, 0, &H20, 0, 0, 0, &HFF, &HFF, 0, 0, &HFF, &HFF, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Generate_WithBinaryResource_WritesResource()
    Dim Writer As ResourceWriter
    Dim Output As New MemoryStream
    Set Writer = Cor.NewResourceWriter(Output)
    
    Writer.AddResource Encoding.ASCII.GetBytes("hello"), 101
    Writer.Generate
    
    Dim Actual() As Byte
    Actual = Output.ToArray
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&H0, &H0, &H0, &H0, &H20, &H0, &H0, &H0, &HFF, &HFF, &H0, &H0, &HFF, &HFF, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H5, &H0, &H0, &H0, &H2C, &H0, &H0, &H0, &H43, &H0, &H55, &H0, &H53, &H0, &H54, &H0, &H4F, &H0, &H4D, &H0, &H0, &H0, &HFF, &HFF, &H65, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H30, &H10, &H9, &H4, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H68, &H65, &H6C, &H6C, &H6F, &H0, &H0, &H0)
End Sub

Public Sub Generate_WhenGenerated_ThrowsInvalidOperationException()
    Dim Writer As ResourceWriter
    Dim Output As New MemoryStream
    Set Writer = Cor.NewResourceWriter(Output)
    Writer.Generate
    
    On Error Resume Next
    Writer.Generate
    
    AssertInvalidOperationException Err
End Sub

Public Sub AddResource_WhenGenerated_ThrowsInvalidOperationException()
    Dim Writer As ResourceWriter
    Dim Output As New MemoryStream
    Set Writer = Cor.NewResourceWriter(Output)
    Writer.Generate
    
    On Error Resume Next
    Writer.AddResource "hello", 101
    
    AssertInvalidOperationException Err
End Sub
