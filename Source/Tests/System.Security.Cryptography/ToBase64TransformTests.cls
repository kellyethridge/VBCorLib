VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToBase64TransformTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ToBase64TransformTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "TransformBlock_WithInputOffset_TransformSubarray"
            Test.Use 0, Cor.NewBytes(&H41, &H51, &H49, &H44)
            Test.Use 1, Cor.NewBytes(&H41, &H67, &H4D, &H45)
            Test.Use 2, Cor.NewBytes(&H41, &H77, &H51, &H46)
            Test.Use 3, Cor.NewBytes(&H42, &H41, &H55, &H47)
            Test.Use 4, Cor.NewBytes(&H42, &H51, &H59, &H48)
            Test.Use 5, Cor.NewBytes(&H42, &H67, &H63, &H49)
            
        Case "TransformFinalBlock_WithInputOffset_ReturnsExpected"
            Test.Use 0, Cor.NewBytes(&H41, &H51, &H49, &H44)
            Test.Use 1, Cor.NewBytes(&H41, &H67, &H4D, &H45)
            Test.Use 2, Cor.NewBytes(&H41, &H77, &H51, &H46)
            Test.Use 3, Cor.NewBytes(&H42, &H41, &H55, &H47)
            Test.Use 4, Cor.NewBytes(&H42, &H51, &H59, &H48)
            Test.Use 5, Cor.NewBytes(&H42, &H67, &H63, &H49)
            
        Case "TransformFinalBlock_WithLessThan3Input_ReturnsExpected"
            Test.Use 6, Cor.NewBytes(&H42, &H77, &H67, &H3D)
            Test.Use 7, Cor.NewBytes(&H43, &H41, &H3D, &H3D)
            Test.Use 8, Cor.NewBytes()
            
    End Select
End Sub

Public Sub Constructor_WhenCalled_SetsDefaultValues()
    Dim Trans As New ToBase64Transform
    
    Assert.That Trans.CanReuseTransform, Iz.True
    Assert.That Trans.CanTransformMultipleBlocks, Iz.False
    Assert.That Trans.InputBlockSize, Iz.EqualTo(3)
    Assert.That Trans.OutputBlockSize, Iz.EqualTo(4)
End Sub

Public Sub TransformBlock_WithInputOffset_TransformSubarray(ByVal InputOffset As Long, ByRef Expected As Variant)
    Dim Trans As New ToBase64Transform
    Dim InputBuffer() As Byte
    InputBuffer = Cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Dim Actual(3) As Byte
    Trans.TransformBlock InputBuffer, InputOffset, 3, Actual, 0
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformBlock_WithNullInputBuffer_ThrowsArgumentNullException()
    Dim Trans As New ToBase64Transform
    Dim NullBytes() As Byte
    Dim OutputBuffer(3) As Byte
    
    On Error Resume Next
    Trans.TransformBlock NullBytes, 0, 3, OutputBuffer, 0
    
    AssertArgumentNullException Err, "InputBuffer"
End Sub

Public Sub TransformBlock_WithInputOffsetLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Trans As New ToBase64Transform
    Dim InputBuffer(2 To 5) As Byte
    Dim OutputBuffer(3) As Byte
    
    On Error Resume Next
    Trans.TransformBlock InputBuffer, 1, 3, OutputBuffer, 0
    
    AssertArgumentOutOfRangeException Err, "InputOffset"
End Sub

Public Sub TransformBlock_WithNegativeInputCount_ThrowsArgumentOutOfRangeException()
    Dim Trans As New ToBase64Transform
    Dim InputBuffer(0 To 5) As Byte
    Dim OutputBuffer(3) As Byte
    
    On Error Resume Next
    Trans.TransformBlock InputBuffer, 0, -1, OutputBuffer, 0
    
    AssertArgumentOutOfRangeException Err, "InputCount"
End Sub

Public Sub TransformBlock_WithTooSmallInputBuffer_ThrowsArgumentException()
    Dim Trans As New ToBase64Transform
    Dim InputBuffer(0 To 1) As Byte
    Dim OutputBuffer(3) As Byte
    
    On Error Resume Next
    Trans.TransformBlock InputBuffer, 0, 0, OutputBuffer, 0
    
    AssertArgumentException Err
End Sub

Public Sub TransformBlock_WithNullOutputBuffer_ThrowsArgumentNullException()
    Dim Trans As New ToBase64Transform
    Dim NullBytes() As Byte
    Dim InputBuffer(3) As Byte
    
    On Error Resume Next
    Trans.TransformBlock InputBuffer, 0, 3, NullBytes, 0
    
    AssertArgumentNullException Err, "OutputBuffer"
End Sub

Public Sub TransformBlock_WithOutputOffsetLessThanLB_ThrowsArgumentOutOfRangeException()
    Dim Trans As New ToBase64Transform
    Dim InputBuffer(2 To 5) As Byte
    Dim OutputBuffer(3) As Byte
    
    On Error Resume Next
    Trans.TransformBlock InputBuffer, 2, 3, OutputBuffer, -1
    
    AssertArgumentOutOfRangeException Err, "OutputOffset"
End Sub

Public Sub TransformBlock_WithOuputOffsetCloserThan4ToUbound_ThrowsArgumentException()
    Dim Trans As New ToBase64Transform
    Dim InputBuffer(0 To 5) As Byte
    Dim OutputBuffer(3) As Byte
    
    On Error Resume Next
    Trans.TransformBlock InputBuffer, 1, 3, OutputBuffer, 1
    
    AssertArgumentException Err, "OutputOffset"
End Sub

Public Sub TransformFinalBlock_WithInputOffset_ReturnsExpected(ByVal InputOffset As Long, ByRef Expected As Variant)
    Dim Trans As New ToBase64Transform
    Dim InputBuffer() As Byte
    InputBuffer = Cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(InputBuffer, InputOffset, 3)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithLessThan3Input_ReturnsExpected(ByVal InputOffset As Long, ByRef Expected As Variant)
    Dim Trans As New ToBase64Transform
    Dim InputBuffer() As Byte
    InputBuffer = Cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(InputBuffer, InputOffset, 8 - InputOffset)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithNullInputBuffer_ThrowsArgumentNullException()
    Dim Trans As New ToBase64Transform
    Dim NullBytes() As Byte
    
    On Error Resume Next
    Trans.TransformFinalBlock NullBytes, 0, 3
    
    AssertArgumentNullException Err, "InputBuffer"
End Sub

Public Sub TransformFinalBlock_WithInputOffsetLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Trans As New ToBase64Transform
    Dim InputBuffer(2 To 5) As Byte
    
    On Error Resume Next
    Trans.TransformFinalBlock InputBuffer, 1, 3
    
    AssertArgumentOutOfRangeException Err, "InputOffset"
End Sub

Public Sub TransformFinalBlock_WithNegativeInputCount_ThrowsArgumentOutOfRangeException()
    Dim Trans As New ToBase64Transform
    Dim InputBuffer(0 To 5) As Byte
    
    On Error Resume Next
    Trans.TransformFinalBlock InputBuffer, 0, -1
    
    AssertArgumentOutOfRangeException Err, "InputCount"
End Sub


