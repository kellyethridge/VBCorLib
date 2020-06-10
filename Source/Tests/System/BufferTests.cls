VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BufferTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BufferTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ByteLength_WithValues_ReturnsExpected"
            Dim Bytes() As Byte
            Dim Longs() As Long

            ReDim Bytes(0 To 9)
            Test.Use Bytes, 10
            ReDim Bytes(1 To 4, 1 To 4)
            Test.Use Bytes, 16

            ReDim Longs(1 To 10)
            Test.Use Longs, 40
            ReDim Longs(1 To 10, 1 To 10)
            Test.Use Longs, 400

    End Select
End Sub

Public Sub BlockCopy_WithByteArrays_CopiesBytes()
    Dim dst(0 To 2) As Byte
    
    Buffer.BlockCopy NewBytes(1, 2, 3), 0, dst, 0, 3
    
    Assert.That dst, Iz.EqualTo(NewBytes(1, 2, 3))
End Sub

Public Sub BlockCopy_WithLongToByteArray_CopiesToBytes()
    Dim dst(0 To 3) As Byte
    
    Buffer.BlockCopy NewLongs(&H12345678), 0, dst, 0, 4
    
    Assert.That dst, Iz.EqualTo(NewBytes(&H78, &H56, &H34, &H12))
End Sub

Public Sub BlockCopy_WithSubsetToSubset_CopiesBytes()
    Dim dst(0 To 3) As Byte
    
    Buffer.BlockCopy NewLongs(&H12345678), 0, dst, 2, 2
    
    Assert.That dst, Iz.EqualTo(NewBytes(0, 0, &H78, &H56))
End Sub

Public Sub BlockCopy_WithSubsetToSubset_CopiesBytes2()
    Dim dst(0 To 3) As Byte
    
    Buffer.BlockCopy NewLongs(&H12345678), 2, dst, 0, 2
    
    Assert.That dst, Iz.EqualTo(NewBytes(&H34, &H12, 0, 0))
End Sub

Public Sub BlockCopy_WithMultiDimToBytes_CopiesBytes()
    Dim src(1, 2) As Byte
    Dim dst(5) As Byte
    src(0, 0) = 1
    src(1, 0) = 2
    src(0, 1) = 3
    src(1, 1) = 4
    src(0, 2) = 5
    src(1, 2) = 6
    
    Buffer.BlockCopy src, 0, dst, 0, 6
    
    Assert.That dst, Iz.EqualTo(NewBytes(1, 2, 3, 4, 5, 6))
End Sub

Public Sub ByteLength_WithValues_ReturnsExpected(ByVal Arr As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = Buffer.ByteLength(Arr)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BlockCopy_WithNullSource_ThrowsArgumentNullException()
    Dim NullSource() As Byte
    Dim Target(0 To 10) As Byte
    
    On Error Resume Next
    Buffer.BlockCopy NullSource, 0, Target, 0, 1
    
    AssertArgumentNullException Err, "Source"
End Sub

Public Sub BlockCopy_WithNullDestination_ThrowsArgumentNullException()
    Dim Source(0 To 10) As Byte
    Dim NullDestination() As Byte
    
    On Error Resume Next
    Buffer.BlockCopy Source, 0, NullDestination, 0, 1
    
    AssertArgumentNullException Err, "Destination"
End Sub

Public Sub BlockCopy_WithSourceIndexLessThanLowerBound_ThrowsArgumentOutOfRangeException()
    Dim src(0 To 10) As Byte
    Dim dst(0 To 10) As Byte
    
    On Error Resume Next
    Buffer.BlockCopy src, -1, dst, 0, 1
    
    AssertArgumentOutOfRangeException Err, "SourceOffset"
End Sub

Public Sub BlockCopy_WithDestinationIndexLessThanLowerBound_ThrowsArgumentOutOfRangeException()
    Dim src(0 To 10) As Byte
    Dim dst(0 To 10) As Byte
    
    On Error Resume Next
    Buffer.BlockCopy src, 0, dst, -1, 1
    
    AssertArgumentOutOfRangeException Err, "DestinationOffset"
End Sub

Public Sub BlockCopy_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    Dim src(0 To 10) As Byte
    Dim dst(0 To 10) As Byte
    
    On Error Resume Next
    Buffer.BlockCopy src, 0, dst, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub BlockCopy_WithCountExtendingPassedEndOfSource_ThrowsArgumentException()
    Dim src(0) As Long
    Dim dst(0) As Long
    
    On Error Resume Next
    Buffer.BlockCopy src, 1, dst, 0, 4
    
    AssertArgumentException Err
End Sub

Public Sub BlockCopy_WithCountExtendingPassedEndOfDestination_ThrowsArgumentException()
    Dim src(0) As Long
    Dim dst(0) As Long
    
    On Error Resume Next
    Buffer.BlockCopy src, 0, dst, 1, 4
    
    AssertArgumentException Err
End Sub

Public Sub GetByte_WhenCalled_ReturnsByteValue()
    Dim Actual As Byte
    
    Actual = Buffer.GetByte(NewBytes(1, 2, 3), 1)
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub GetByte_WithNullArray_ThrowsArgumentNullException()
    Dim Arr() As Byte
    
    On Error Resume Next
    Buffer.GetByte Arr, 0
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub GetByte_WithNegativeIndex_ThrowsArgumentOutOfRangeExceptioin()
    Dim Arr(0) As Byte
    
    On Error Resume Next
    Buffer.GetByte Arr, -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetByte_WithIndexGreaterThanTheArrayLength_ThrowsArgumentOutOfRangeException()
    Dim Arr(1 To 10) As Byte
    
    On Error Resume Next
    Buffer.GetByte Arr, 10
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub SetByte_WithValue_SetsByteInArray()
    Dim Arr(0) As Byte
    
    Buffer.SetByte Arr, 0, 1
    
    Assert.That Arr(0), Iz.EqualTo(1)
End Sub

