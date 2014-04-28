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
            
        Case "ReadByte_WithSeekFromBeginning_ReturnsExpected"
            Test.Use 0, 1
            Test.Use 4, 5
            Test.Use 5, -1
            
        Case "SeekPosition_WhenFromCurrent_SetsPosition"
            Test.Use 0, 2
            Test.Use -2, 0
            Test.Use 2, 4
            Test.Use 3, 5
            
        Case "ReadByte_WithSeekFromCurrent_ReturnsExpected"
            Test.Use 0, 3
            Test.Use -2, 1
            Test.Use 2, 5
            Test.Use 3, -1
            
        Case "SeekPosition_WhenFromEnd_SetsPosition"
            Test.Use 0, 5
            Test.Use -1, 4
            Test.Use -5, 0
            
        Case "ReadByte_WithSeekFromEnd_ReturnsExpected"
            Test.Use 0, -1
            Test.Use -1, 5
            Test.Use -5, 1
            
    End Select
End Sub

Public Sub ReadByte_WithGivenPosition_ReturnsExpected(ByVal Position As Long, ByVal Expected As Long)
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source)
    ms.Position = Position
    
    Dim Actual As Long
    Actual = ms.ReadByte
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadByte_WithGivenPositionInSubset_ReturnsExpected(ByVal Position As Long, ByVal Expected As Long)
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source, 1, 3)
    ms.Position = Position
    
    Dim Actual As Long
    Actual = ms.ReadByte
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SeekPosition_WhenFromBeginning_SetsPosition(ByVal Offset As Long, ByVal Expected As Long)
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source)
    
    Dim Actual As Long
    Actual = ms.SeekPosition(Offset, SeekOrigin.FromBeginning)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadByte_WithSeekFromBeginning_ReturnsExpected(ByVal Offset As Long, ByVal Expected As Long)
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source)
    ms.SeekPosition Offset, SeekOrigin.FromBeginning
    
    Dim Actual As Long
    Actual = ms.ReadByte
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SeekPosition_WhenFromCurrent_SetsPosition(ByVal Offset As Long, ByVal Expected As Long)
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source)
    ms.Position = 2
    
    Dim Actual As Long
    Actual = ms.SeekPosition(Offset, SeekOrigin.FromCurrent)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadByte_WithSeekFromCurrent_ReturnsExpected(ByVal Offset As Long, ByVal Expected As Long)
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source)
    ms.Position = 2
    ms.SeekPosition Offset, SeekOrigin.FromCurrent
    
    Dim Actual As Long
    Actual = ms.ReadByte
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SeekPosition_WhenFromEnd_SetsPosition(ByVal Offset As Long, ByVal Expected As Long)
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source)
    ms.Position = 2
    
    Dim Actual As Long
    Actual = ms.SeekPosition(Offset, SeekOrigin.FromEnd)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadByte_WithSeekFromEnd_ReturnsExpected(ByVal Offset As Long, ByVal Expected As Long)
    Dim Source() As Byte
    Source = NewBytes(1, 2, 3, 4, 5)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Source)
    ms.SeekPosition Offset, SeekOrigin.FromEnd
    
    Dim Actual As Long
    Actual = ms.ReadByte
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub SeekPosition_WithInvalidSeekOrigin_ThrowsArgumentException()
    Dim ms As New MemoryStream
    
    On Error Resume Next
    ms.SeekPosition 0, 5
    
    AssertArgumentException Err
End Sub




'Implements AsyncCallback
'
'Private ms As MemoryStream
'Private mAsyncCallback As Boolean
'Private b() As Byte
'
'Public Sub Setup()
'    Set ms = New MemoryStream
'    mAsyncCallback = False
'End Sub
'
'Public Sub Position_WhenNew_EqualsZero()
'    Dim Stream As New MemoryStream
'
'    Dim Actual As Long
'    Actual = Stream.Position
'
'    Assert.That Actual, Iz.EqualTo(0)
'End Sub
'
'Public Sub Length_WhenNew_EqualsZero()
'    Dim Stream As New MemoryStream
'
'    Dim Actual As Long
'    Actual = Stream.Length
'
'    Assert.That Actual, Iz.EqualTo(0)
'End Sub
'
'Public Sub CanWrite_WhenNew_ReturnsTrue()
'    Dim Stream As New MemoryStream
'
'    Dim Actual As Boolean
'    Actual = Stream.CanWrite
'
'    Assert.That Actual, Iz.True
'End Sub
'
'Public Sub WriteByte_WhenCalled_AddsToStream()
'    Dim StreamData(10) As Byte
'    Dim Stream As MemoryStream
'    Set Stream = Cor.NewMemoryStream(StreamData)
'
'    Stream.WriteByte 1
'
'    Assert.That Stream.Position, Iz.EqualTo(1)
'    Assert.That StreamData(0), Iz.EqualTo(1)
'End Sub
'
'Public Sub ReadByte_WhenPositionSet_ReturnsExpected()
'    Dim Source() As Byte
'    Source = NewBytes(1, 2, 3, 4, 5)
'    Dim ms As MemoryStream
'    Set ms = Cor.NewMemoryStream(Source)
'    ms.Position = 2
'
'    Dim Actual As Byte
'    Actual = ms.ReadByte
'
'    Assert.That Actual, Iz.EqualTo(3)
'End Sub
'
'Public Sub TestReadPreviousByte()
'    ms.WriteByte 55
'    ms.Position = 0
'    Assert.AreEqual 55, ms.ReadByte
'End Sub
'
'Public Sub TestReadByteAtEndOfStream()
'    Assert.AreEqual -1, ms.ReadByte
'End Sub
'
'Public Sub TestWriteBlock()
'    Dim i As Long
'
'    ReDim b(9)
'    For i = 0 To 9
'        b(i) = i + 1
'    Next i
'    ms.WriteBlock b, 0, 10
'    Assert.AreEqual 10@, ms.Position
'End Sub
'
'Public Sub TestReadBlock()
'    Dim i As Long
'
'    ReDim b(10)
'    For i = 0 To 9
'        b(i) = i + 1
'    Next i
'    b(10) = 13
'    For i = 1 To 1000
'        ms.WriteBlock b, 0, 11
'    Next i
'    ms.Position = 0
'
'    Dim b2(10) As Byte
'    Dim j As Long
'
'    For j = 1 To 1000
'        Assert.AreEqual ms.ReadBlock(b2, 0, 11), 11
'        For i = 0 To 9
'            Assert.AreEqual i + 1, b2(i)
'        Next i
'        Assert.AreEqual 13, b2(10)
'    Next j
'End Sub
'
'Public Sub TestBeginWrite()
'    Dim i As Long
'    Dim ar As IAsyncResult
'
'    ReDim b(9)
'    For i = 0 To 9
'        b(i) = i + 1
'    Next i
'
'    Set ar = ms.BeginWrite(b, 0, 10, Me)
'    Assert.IsNotNothing ar, "ar should not be Nothing."
'    Assert.IsTrue mAsyncCallback, "AsyncCallback should be executed."
'    ms.EndWrite ar
'    Assert.AreEqual 10@, ms.Length
'End Sub
'
'Public Sub TestBeginRead()
'    Dim i As Long
'    Dim ar As IAsyncResult
'
'    ReDim b(9)
'    For i = 0 To 9
'        b(i) = i + 1
'    Next i
'    ms.WriteBlock b, 0, 10
'
'    ReDim b(9)
'    ms.SeekPosition 0, FromBeginning
'    Set ar = ms.BeginRead(b, 0, 10, Me)
'    Assert.IsNotNothing ar, "ar should not be Nothing."
'    Assert.IsTrue mAsyncCallback, "AsyncCallback should be executed."
'    Assert.AreEqual 10, ms.EndRead(ar)
'    For i = 0 To 9
'        Assert.AreEqual i + 1, b(i)
'    Next i
'End Sub
'
'Public Sub TestReadToOneBytePassedEnd()
'    Dim i As Long
'    Dim b(10) As Byte
'
'    For i = 1 To 10
'        ms.WriteByte i
'    Next i
'
'    ms.Position = 8
'    Assert.AreEqual 9, ms.ReadByte
'    ms.Position = 8
'    Assert.AreEqual 2, ms.ReadBlock(b, 0, 3)
'    Assert.AreEqual 9, b(0)
'    Assert.AreEqual 10, b(1)
'End Sub
'
'
'
'Private Sub AsyncCallback_Execute(ByVal ar As VBCorLib.IAsyncResult)
'    mAsyncCallback = True
'End Sub

