VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryStreamTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMemoryStream
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ReadByte_WithGivenPosition_ReturnsExpected"
            Test.Use 0, 1
            Test.Use 2, 3
            Test.Use 4, 5
            Test.Use 5, -1
            
        Case "ReadByte_WithGivenPositionInSubset_ReturnsExpected"
            Test.Use 0, 2
            Test.Use 2, 4
            Test.Use 3, -1
            
        Case "SeekPosition_WhenFromBeginning_SetsPosition"
            Test.Use 0, 0
            Test.Use 4, 4
            Test.Use 5, 5
            
        Case "SeekPosition_WhenFromCurrent_SetsPosition"
            Test.Use 0, 2
            Test.Use -2, 0
            Test.Use 2, 4
            Test.Use 3, 5
            
        Case "SeekPosition_WhenFromEnd_SetsPosition"
            Test.Use 0, 5
            Test.Use -1, 4
            Test.Use -5, 0
            
        Case "SeekPosition_WhenFromBeginningInSubset_SetsPosition"
            Test.Use 0, 0
            Test.Use 2, 2
            Test.Use 3, 3
            
        Case "SeekPosition_WhenFromCurrentInSubset_SetsPosition"
            Test.Use 0, 1
            Test.Use -1, 0
            Test.Use 1, 2
            Test.Use 2, 3
        
        Case "SeekPosition_WhenFromEndInSubset_SetsPosition"
            Test.Use 0, 3
            Test.Use -1, 2
            Test.Use -3, 0
            
    End Select
End Sub

Public Sub ReadByte_WithGivenPosition_ReturnsExpected(ByVal Position As Long, ByVal Expected As Long)
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    ms.Position = Position
    
    Dim Actual As Long
    Actual = ms.ReadByte
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadByte_WhenCalled_AdvancesPosition()
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    
    ms.ReadByte
    
    Assert.That ms.Position, Iz.EqualTo(1)
End Sub

Public Sub ReadByte_WithGivenPositionInSubset_ReturnsExpected(ByVal Position As Long, ByVal Expected As Long)
    Dim ms As MemoryStream
    Set ms = MakeMemoryStreamSubset
    ms.Position = Position
    
    Dim Actual As Long
    Actual = ms.ReadByte
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteByte_WhenCalled_WritesToBuffer()
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source)
    
    ms.WriteByte 10
    
    Assert.That Source, Iz.EqualTo(NewBytes(10, 2, 3, 4, 5))
End Sub

Public Sub WriteByte_WhenCalled_AdvancesPosition()
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source)
    
    ms.WriteByte 10

    Assert.That ms.Position, Iz.EqualTo(1)
End Sub

Public Sub WriteByte_WithSubset_WritesToBuffer()
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source, 1, 3)
    
    ms.WriteByte 10
    
    Assert.That Source, Iz.EqualTo(NewBytes(1, 10, 3, 4, 5))
End Sub

Public Sub SeekPosition_WhenFromBeginning_SetsPosition(ByVal Offset As Long, ByVal Expected As Long)
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    
    Dim Actual As Long
    Actual = ms.SeekPosition(Offset, SeekOrigin.FromBeginning)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SeekPosition_WhenFromCurrent_SetsPosition(ByVal Offset As Long, ByVal Expected As Long)
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    ms.Position = 2
    
    Dim Actual As Long
    Actual = ms.SeekPosition(Offset, SeekOrigin.FromCurrent)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SeekPosition_WhenFromEnd_SetsPosition(ByVal Offset As Long, ByVal Expected As Long)
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    
    Dim Actual As Long
    Actual = ms.SeekPosition(Offset, SeekOrigin.FromEnd)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SeekPosition_WithInvalidSeekOrigin_ThrowsArgumentException()
    Const InvalidOrigin As Long = 5
    Dim ms As New MemoryStream
    
    On Error Resume Next
    ms.SeekPosition 0, InvalidOrigin
    
    AssertArgumentException Err
End Sub

Public Sub SeekPosition_WhenFromBeginningInSubset_SetsPosition(ByVal Offset As Long, ByVal Expected As Long)
    Dim ms As MemoryStream
    Set ms = MakeMemoryStreamSubset
    
    Dim Actual As Long
    Actual = ms.SeekPosition(Offset, SeekOrigin.FromBeginning)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SeekPosition_WhenFromCurrentInSubset_SetsPosition(ByVal Offset As Long, ByVal Expected As Long)
    Dim ms As MemoryStream
    Set ms = MakeMemoryStreamSubset
    ms.Position = 1
    
    Dim Actual As Long
    Actual = ms.SeekPosition(Offset, SeekOrigin.FromCurrent)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SeekPosition_WhenFromEndInSubset_SetsPosition(ByVal Offset As Long, ByVal Expected As Long)
    Dim ms As MemoryStream
    Set ms = MakeMemoryStreamSubset
    
    Dim Actual As Long
    Actual = ms.SeekPosition(Offset, SeekOrigin.FromEnd)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Constructor_WithSharedMode_SharesBuffer()
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    
    Set ms = Cor.NewMemoryStream(Source, Mode:=ShareMode)
    
    Source(0) = 10
    Assert.That ms.ReadByte, Iz.EqualTo(10)
End Sub

Public Sub Constructor_WithCopyMode_CopiesBuffer()
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    
    Set ms = Cor.NewMemoryStream(Source, Mode:=CopyMode)
    
    Source(0) = 10
    Assert.That ms.ReadByte, Iz.EqualTo(1)
End Sub

Public Sub Constructor_WithCopyMode_DoesNotChangeSourceBuffer()
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    
    Set ms = Cor.NewMemoryStream(Source, Mode:=CopyMode)
        
    Assert.That Source, Iz.EqualTo(NewBytes(1, 2, 3, 4, 5))
End Sub

Public Sub Constructor_WithAcquireMode_AcquiredBuffer()
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    
    Set ms = Cor.NewMemoryStream(Source, Mode:=AcquireMode)
    
    Assert.That ms.ReadByte, Iz.EqualTo(1)
End Sub

Public Sub Constructor_WithAcquireMode_NullsSourceBuffer()
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    
    Set ms = Cor.NewMemoryStream(Source, Mode:=AcquireMode)
    
    Assert.That cArray.IsNull(Source), Iz.True
End Sub

Public Sub ReadBlock_WhenCalled_ReturnsAmountRead()
    Dim OutputBuffer(0 To 2) As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
        
    Dim Actual As Long
    Actual = ms.ReadBlock(OutputBuffer, 0, 3)
    
    Assert.That Actual, Iz.EqualTo(3)
End Sub

Public Sub ReadBlock_WithOffsetOfZero_FillsStartOfOutputBuffer()
    Dim OutputBuffer(0 To 4) As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
        
    ms.ReadBlock OutputBuffer, 0, 3

    Assert.That OutputBuffer, Iz.EqualTo(NewBytes(1, 2, 3, 0, 0))
End Sub

Public Sub ReadBlock_WithOffsetOfTwo_FillsEndOfOutputBuffer()
    Dim OutputBuffer(0 To 4) As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
        
    ms.ReadBlock OutputBuffer, 2, 3

    Assert.That OutputBuffer, Iz.EqualTo(NewBytes(0, 0, 1, 2, 3))
End Sub

Public Sub ReadBlock_WithOffsetAndNonZeroLowerBoundArray_FillsEndOfOutputBuffer()
    Dim OutputBuffer(2 To 6) As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
        
    ms.ReadBlock OutputBuffer, 2, 3

    Assert.That OutputBuffer, Iz.EqualTo(NewBytes(0, 0, 1, 2, 3)).AsCollection
End Sub

Public Sub ReadBlock_WithNegativeOffset_ThrowsArgumentOutOfRangeException()
    Dim OutputBuffer(0 To 4) As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
        
    On Error Resume Next
    ms.ReadBlock OutputBuffer, -1, 3
    
    AssertArgumentOutOfRangeException Err, "Offset"
End Sub

Public Sub ReadBlock_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    Dim OutputBuffer(0 To 4) As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
        
    On Error Resume Next
    ms.ReadBlock OutputBuffer, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub ReadBlock_WithNullOutputBuffer_ThrowsArgumentNullException()
    Dim OutputBuffer() As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
        
    On Error Resume Next
    ms.ReadBlock OutputBuffer, 0, 3
    
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub ReadBlock_WithMultipleDimensionOutputBuffer_ThrowsRankException()
    Dim OutputBuffer(0 To 3, 0 To 3) As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
        
    On Error Resume Next
    ms.ReadBlock OutputBuffer, 0, 3
    
    AssertRankException Err
End Sub

Public Sub WriteBlock_WhenCalled_StoresInBuffer()
    Dim Buffer(0 To 4) As Byte
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Buffer)
    
    ms.WriteBlock NewBytes(7, 8, 9), 0, 3
    
    Assert.That Buffer, Iz.EqualTo(NewBytes(7, 8, 9, 0, 0))
End Sub

Public Sub WriteBlock_WithOffsetOfTwo_StoresEndOfSourceToBuffer()
    Dim Buffer(0 To 4) As Byte
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Buffer)
    
    ms.WriteBlock NewBytes(7, 8, 9, 10, 11), 2, 3
    
    Assert.That Buffer, Iz.EqualTo(NewBytes(9, 10, 11, 0, 0))
End Sub

Public Sub WriteBlock_WithNegativeOffset_ThrowsArgumentOutOfRangeException()
    Dim Buffer(0 To 4) As Byte
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Buffer)
    
    On Error Resume Next
    ms.WriteBlock NewBytes(7, 8, 9, 10, 11), -1, 3
    
    AssertArgumentOutOfRangeException Err, "Offset"
End Sub

Public Sub WriteBlock_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    Dim Buffer(0 To 4) As Byte
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Buffer)
    
    On Error Resume Next
    ms.WriteBlock NewBytes(7, 8, 9, 10, 11), 2, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub WriteBlock_WithNullArray_ThrowsArgumentNullException()
    Dim ms As New MemoryStream
    
    On Error Resume Next
    Dim NullSource() As Byte
    ms.WriteBlock NullSource, 2, 3

    AssertArgumentNullException Err, "Buffer"
End Sub

Public Sub WriteBlock_WithMultiDimensionArray_ThrowsRankException()
    Dim ms As New MemoryStream
    
    On Error Resume Next
    Dim MultiDimArray(0 To 2, 0 To 2) As Byte
    ms.WriteBlock MultiDimArray, 2, 3

    AssertRankException Err
End Sub

Public Sub WriteBlock_WithInvalidRange_ThrowsArgumentException()
    Dim ms As New MemoryStream
    
    On Error Resume Next
    ms.WriteBlock NewBytes(1, 2, 3), 1, 3
    
    AssertArgumentException Err
End Sub

Public Sub ToArray_WithNoWrittenContent_ReturnsEmptyArray()
    Dim ms As New MemoryStream
    
    Dim Actual() As Byte
    Actual = ms.ToArray
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub ToArray_WithContentWritten_ReturnsArrayOfContent()
    Dim ms As New MemoryStream
    ms.WriteBlock NewBytes(1, 2, 3), 0, 3
    
    Dim Actual() As Byte
    Actual = ms.ToArray
    
    Assert.That Actual, Iz.EqualTo(NewBytes(1, 2, 3))
End Sub

Public Sub Capacity_WithNegativeValue_ThrowsArgumentOutOfRangeException()
    Dim ms As New MemoryStream
    
    On Error Resume Next
    ms.Capacity = -1
    
    AssertArgumentOutOfRangeException Err, "Value"
End Sub

Public Sub Capacity_WhenSetLessThanStreamLength_ThrowsArgumentOutOfRangeException()
    Dim ms As New MemoryStream
    ms.WriteBlock NewBytes(1, 2, 3, 4, 5), 0, 5
    
    On Error Resume Next
    ms.Capacity = 4
    
    AssertArgumentOutOfRangeException Err, "Value"
End Sub

Public Sub Capacity_WithSharedBuffer_ThrowsNotSupportedException()
    Dim Source(0 To 10) As Byte
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source, Mode:=ShareMode)
    
    On Error Resume Next
    ms.Capacity = ms.Capacity + 1
    
    AssertNotSupportedException Err
End Sub

Public Sub Capacity_WithInternalBuffer_ExpandsStream()
    Dim ms As New MemoryStream
    
    ms.Capacity = ms.Capacity + 1
End Sub

Public Sub Capacity_WithCopiedBuffer_ExpandsStream()
    Dim Source(0 To 10) As Byte
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source, Mode:=CopyMode)
    
    ms.Capacity = ms.Capacity + 1
End Sub

Public Sub Capacity_WithAcquiredBuffer_ExpandsStream()
    Dim Source() As Byte
    ReDim Source(0 To 10)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source, Mode:=AcquireMode)
    
    ms.Capacity = ms.Capacity + 1
End Sub

Public Sub WriteByte_WithUnwritableStream_ThrowsNotSupportedException()
    Dim Source(0 To 10) As Byte
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source, Writable:=False)
    
    On Error Resume Next
    ms.WriteByte 1
    
    AssertNotSupportedException Err
End Sub

Public Sub WriteByte_WithClosedStream_ThrowsObjectDisposedException()
    Dim ms As New MemoryStream
    ms.CloseStream
    
    On Error Resume Next
    ms.WriteByte 1
    
    AssertObjectDisposedException Err
End Sub

Public Sub WriteTo_WithNothing_ThrowsArgumentNullException()
    Dim ms As New MemoryStream
    
    On Error Resume Next
    ms.WriteTo Nothing
    
    AssertArgumentNullException Err, "Stream"
End Sub

Public Sub WriteTo_WhenCalled_WritesContentsToTargetStream()
    Dim ms As New MemoryStream
    ms.WriteBlock NewBytes(1, 2, 3), 0, 3
    
    Dim Target As New MemoryStream
    ms.WriteTo Target
    
    Assert.That Target.ToArray, Iz.EqualTo(NewBytes(1, 2, 3))
End Sub

Public Sub BeginRead_WhenCalled_ReturnsAsyncResult()
    Dim Buffer(0 To 9) As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    
    Dim Result As IAsyncResult
    Set Result = ms.BeginRead(Buffer, 0, 5)
    
    Assert.That Result, Iz.Not.Nothing
    Assert.That Result.IsCompleted, Iz.True
    Assert.That Buffer, Iz.EqualTo(NewBytes(1, 2, 3, 4, 5, 0, 0, 0, 0, 0))
End Sub

Public Sub EndRead_WhenCalled_ReturnsCountRead()
    Dim Buffer(0 To 9) As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    Dim Result As IAsyncResult
    Set Result = ms.BeginRead(Buffer, 0, 5)
    
    Dim Actual As Long
    Actual = ms.EndRead(Result)
    
    Assert.That Actual, Iz.EqualTo(5)
End Sub

Public Sub BeginRead_WithAsyncCallback_ExecutesCallback()
    Dim Buffer(0 To 9) As Byte
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    Dim Callback As New FakeAsyncCallback
    
    Dim Result As IAsyncResult
    Set Result = ms.BeginRead(Buffer, 0, 5, Callback)

    Assert.That Callback.Called, Iz.True
End Sub

Public Sub SetLength_WhenLessThanCurrentLength_TruncatesStream()
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    
    ms.SetLength 3
    
    Assert.That ms.ToArray, Iz.EqualTo(NewBytes(1, 2, 3))
End Sub

Public Sub SetLength_WhenGreaterThanCurrentLength_FillsWithZeros()
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    
    ms.SetLength 8
    
    Assert.That ms.ToArray, Iz.EqualTo(NewBytes(1, 2, 3, 4, 5, 0, 0, 0))
End Sub

Public Sub SetLength_WithTrucateThenExpand_OverwritesContentsWithZeros()
    Dim ms As MemoryStream
    Set ms = MakeMemoryStream
    
    ms.SetLength 2
    ms.SetLength 5
    
    Assert.That ms.ToArray, Iz.EqualTo(NewBytes(1, 2, 0, 0, 0))
End Sub



Private Function MakeMemoryStream() As MemoryStream
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Set MakeMemoryStream = Cor.NewMemoryStream(Source, Mode:=AcquireMode)
End Function

Private Function MakeMemoryStreamSubset() As MemoryStream
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Set MakeMemoryStreamSubset = Cor.NewMemoryStream(Source, 1, 3, Mode:=AcquireMode)
End Function

Public Sub BeginWrite_WhenCalled_ReturnsAsyncResult()
    Dim ms As New MemoryStream
    
    Dim Result As IAsyncResult
    Set Result = ms.BeginWrite(NewBytes(1, 2, 3), 0, 3)
    
    Assert.That Result, Iz.Not.Nothing
    Assert.That Result.IsCompleted, Iz.True
    Assert.That ms.ToArray, Iz.EqualTo(NewBytes(1, 2, 3))
End Sub

Public Sub EndWrite_WithAsyncCallback_ExecutesCallback()
    Dim ms As New MemoryStream
    Dim Callback As New FakeAsyncCallback
    
    Dim Result As IAsyncResult
    Set Result = ms.BeginWrite(NewBytes(1, 2, 3), 0, 3, Callback)

    Assert.That Callback.Called, Iz.True
End Sub
