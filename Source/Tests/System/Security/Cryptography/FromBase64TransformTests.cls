VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FromBase64TransformTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' FromBase64TransformTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "TransformFinalBlock_WithValues_ReturnsExpected"
            Test.Use NewBytes(114, 108, 112, 55, 81, 115, 110, 69), 0, 8, NewBytes(174, 90, 123, 66, 201, 196)
            Test.Use NewBytes(114, 108, 112, 55, 81, 115, 61, 61), 0, 8, NewBytes(174, 90, 123, 66)
            Test.Use NewBytes(114, 108, 112, 55, 81, 115, 61, 61), 1, 5, NewBytes(150, 158, 208)
                       
    End Select
End Sub

Public Sub Constructors_WhenCalled_SetsDefaultValues()
    Dim Trans As New FromBase64Transform
    
    Assert.That Trans.CanReuseTransform, Iz.True
    Assert.That Trans.CanTransformMultipleBlocks, Iz.False
    Assert.That Trans.InputBlockSize, Iz.EqualTo(1)
    Assert.That Trans.OutputBlockSize, Iz.EqualTo(3)
End Sub

Public Sub TransformFinalBlock_WithValues_ReturnsExpected(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef Expected() As Byte)
    Dim Trans As New FromBase64Transform
    
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(InputBuffer, InputOffset, InputCount)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithDoNotIgnoreWhitespace_ReturnsExpected()
    Dim Trans As FromBase64Transform
    Dim InputBuffer() As Byte
    Set Trans = NewFromBase64Transform(DoNotIgnoreWhiteSpaces)
    InputBuffer = NewBytes(114, 108, 112, 55, 81, 115, 110, 69)
    
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(InputBuffer, 0, 8)
    
    Dim Expected() As Byte
    Expected = NewBytes(174, 90, 123, 66, 201, 196)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithTab_IgnoresTab()
    Dim Trans As New FromBase64Transform
    Dim InputBuffer() As Byte
    InputBuffer = NewBytes(9, 114, 108, 112, 55, 9, 81, 115, 61, 61, 9)
    
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(InputBuffer, 0, 11)
    
    Dim Expected() As Byte
    Expected = NewBytes(174, 90, 123, 66)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithCR_IgnoresCR()
    Dim Trans As New FromBase64Transform
    Dim InputBuffer() As Byte
    InputBuffer = NewBytes(13, 114, 108, 112, 55, 13, 81, 115, 61, 61, 13)
    
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(InputBuffer, 0, 11)
    
    Dim Expected() As Byte
    Expected = NewBytes(174, 90, 123, 66)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithSpaces_IgnoresSpaces()
    Dim Trans As New FromBase64Transform
    Dim InputBuffer() As Byte
    InputBuffer = NewBytes(32, 114, 108, 112, 55, 32, 81, 115, 61, 61, 32)
    
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(InputBuffer, 0, 11)
    
    Dim Expected() As Byte
    Expected = NewBytes(174, 90, 123, 66)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithUnignoredWhitespace_ThrowsFormatException()
    Dim Trans As New FromBase64Transform
    Dim InputBuffer() As Byte
    InputBuffer = NewBytes(7, 114, 108, 112, 55, 81, 115, 61, 61)
    
    On Error Resume Next
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(InputBuffer, 0, 9)
    
    AssertFormatException Err
End Sub

Public Sub TransformFinalBlock_WhenReusing_ReturnsExpected()
    Dim Trans As New FromBase64Transform
    Dim InputBuffer() As Byte
    InputBuffer = NewBytes(32, 114, 108, 112, 55, 32, 81, 115, 61, 61, 32)
    Dim Expected() As Byte
    Expected = NewBytes(174, 90, 123, 66)
    
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(InputBuffer, 0, 11)
    Assert.That Actual, Iz.EqualTo(Expected)
    Actual = Trans.TransformFinalBlock(InputBuffer, 0, 11)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithInvalidLength_ReturnsEmpty()
    Dim Trans As New FromBase64Transform
    
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(NewBytes(114, 108, 112), 0, 3)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub TransformFinalBlock_WithInvalidData_ReturnsEmpty()
    Dim Trans As New FromBase64Transform
    
    Dim Actual() As Byte
    Actual = Trans.TransformFinalBlock(NewBytes(114, 108, 112, 32), 0, 4)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub TransformBlock_WithNullInputBuffer_ThrowsArgumentNullException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformBlock NullBytes, 0, 0, NewBytes(0), 0
    
    AssertArgumentNullException Err, "InputBuffer"
End Sub

Public Sub TransformBlock_WithInputOffsetLessThanLB_ThrowsArgumentOutOfRangeException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformBlock NewBytes(1, 2, 3, 4), -1, 3, OutputBuffer(5), 0
    
    AssertArgumentOutOfRangeException Err, "InputOffset"
End Sub

Public Sub TransformBlock_WithInputOffsetPlusCountGreaterThanUB_ThrowsArgumentException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformBlock NewBytes(1, 2, 3, 4), 2, 3, OutputBuffer(5), 0
    
    AssertArgumentException Err
End Sub

Public Sub TransformBlock_WithInputCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformBlock NewBytes(1, 2, 3, 4), 0, -1, OutputBuffer(5), 0
    
    AssertArgumentOutOfRangeException Err, "InputCount"
End Sub

Public Sub TransformBlock_WithNullOutpuBuffer_ThrowsArgumentNullException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformBlock OutputBuffer(5), 0, 0, NullBytes, 0
    
    AssertArgumentNullException Err, "OutputBuffer"
End Sub

Public Sub TransfromBlock_WithOutputOffsetLessThanLB_ThrowsArgumentOutOfRangeException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformBlock NewBytes(1, 2, 3, 4), 0, 4, OutputBuffer(10), -1
    
    AssertArgumentOutOfRangeException Err, "OutputOffset"
End Sub

Public Sub TransformBlock_WithOutputOffsetGreaterThanUB_ThrowsArgumentException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformBlock NewBytes(114, 108, 112, 55, 81, 115, 61, 61), 0, 8, OutputBuffer(10), 11
    
    AssertArgumentException Err
End Sub

Public Sub TransformFinalBlock_WithNullInputBuffer_ThrowsArgumentNullException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformFinalBlock NullBytes, 0, 1
    
    AssertArgumentNullException Err, "InputBuffer"
End Sub

Public Sub TransformFinalBlock_WithInputOffsetLessThanLB_ThrowsArgumentOutOfRangeException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformFinalBlock OutputBuffer(10), -1, 10
    
    AssertArgumentOutOfRangeException Err, "InputOffset"
End Sub

Public Sub TransformFinalBlock_WithInputCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformFinalBlock OutputBuffer(10), 0, -1
    
    AssertArgumentOutOfRangeException Err, "InputCount"
End Sub

Public Sub TransformFinalBlock_WithInputOffsetPlusCountGreaterThanUB_ThrowsArgumentException()
    Dim Trans As New FromBase64Transform
    
    On Error Resume Next
    Trans.TransformFinalBlock OutputBuffer(10), 5, 10
    
    AssertArgumentException Err
End Sub


Private Property Get NullBytes() As Byte()
End Property

Private Function OutputBuffer(ByVal Size As Long) As Byte()
    OutputBuffer = CorArray.CreateInstance(vbByte, Size)
End Function























