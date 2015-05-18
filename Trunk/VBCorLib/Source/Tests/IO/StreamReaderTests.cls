VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamReaderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamReader
'
Option Explicit
Implements ICategorizable

Private mIO As IOHelper

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub Setup()
    Set mIO = New IOHelper
End Sub

Public Sub Teardown()
    Set mIO = Nothing
End Sub

Public Sub Read_WhenCalledWithEmptyStream_ReturnsMinusOne()
    Dim Actual As Long
    
    With Cor.NewStreamReader(CreateStream)
        Actual = .Read
    End With
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub Read_WhenCalledWithData_ReturnsData()
    With Cor.NewStreamReader(CreateStream("123"))
        Assert.That .Read, Iz.EqualTo(49)
        Assert.That .Read, Iz.EqualTo(50)
        Assert.That .Read, Iz.EqualTo(51)
        Assert.That .Read, Iz.EqualTo(-1)
    End With
End Sub

Public Sub ReadLine_WhenCalledWithEmptyStream_ReturnsNullString()
    Dim Actual As String
    
    With Cor.NewStreamReader(CreateStream)
        Actual = .ReadLine
    End With
    
    Assert.That cString.IsNull(Actual), Iz.True
End Sub

Public Sub ReadLine_WhenCalledWithEmptyLine_ReturnsEmptyLine()
    Dim line As String
    With Cor.NewStreamReader(CreateStream(vbCrLf))
        line = .ReadLine
        Assert.That cString.IsNull(line), Iz.False
        Assert.That line, Iz.Empty
        Assert.That cString.IsNull(.ReadLine), Iz.True
    End With
End Sub

Public Sub ReadLine_WhenCalledWithData_ReturnsData()
    With Cor.NewStreamReader(CreateStream("1234" & vbCrLf & "6789"))
        Assert.That .ReadLine, Iz.EqualTo("1234")
        Assert.That .ReadLine, Iz.EqualTo("6789")
        Assert.That cString.IsNull(.ReadLine), Iz.True
    End With
End Sub

Public Sub ReadLine_WithLargeData_ReadsData()
    Dim Expected As String
    Expected = Space$(10000)
    
    Dim Actual As String
    With Cor.NewStreamReader(CreateStream(Expected))
        Actual = .ReadLine
    End With
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadBlock_WithCountOfZero_ReturnsNoData()
    Dim Chars() As Integer
    ReDim Chars(100)
    
    Dim Actual As Long
    With Cor.NewStreamReader(CreateStream("1234"))
        Actual = .ReadBlock(Chars, 0, 0)
    End With
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub ReadBlock_WithNullArray_ThrowsArgumentNullException()
    Dim NullArray() As Integer
    
    On Error Resume Next
    With Cor.NewStreamReader(CreateStream("1234"))
        .ReadBlock NullArray, 0, 1
    End With
    
    AssertArgumentNullException Err, "Buffer"
End Sub

Public Sub ReadBlock_WithIndexLessThanBufferLBound_ThrowsArgumentOutOfRangeException()
    Dim Chars(1 To 10) As Integer
    
    On Error Resume Next
    With Cor.NewStreamReader(CreateStream("1234"))
        .ReadBlock Chars, 0, 1
    End With
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub ReadBlock_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Chars(0 To 10) As Integer
    
    On Error Resume Next
    With Cor.NewStreamReader(CreateStream("1234"))
        .ReadBlock Chars, 0, -1
    End With
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub ReadBlock_ThroughBufferSize_ReadsData()
    Dim Chars(0 To 127) As Integer
    Dim Actual As New StringBuilder
    
    With Cor.NewStreamReader(CreateStream(String$(600, 48)), BufferSize:=192)
        Dim CharsRead As Long
        Do
            CharsRead = .ReadBlock(Chars, 0, 128)
            Actual.Append Chars, 0, CharsRead
        Loop While CharsRead > 0
    End With
    
    Assert.That Actual.ToString, Iz.EqualTo(String$(600, 48))
End Sub

Public Sub Peek_WhenCalled_ReturnsCurrentCharacter()
    Dim Actual As Long
    
    With Cor.NewStreamReader(CreateStream("123"))
        Actual = .Peek
    End With
    
    Assert.That Actual, Iz.EqualTo(49)
End Sub

Public Sub Peek_WhenCalled_DoesNotAdvanceStreamPosition()
    Dim Actual As Long
    
    With Cor.NewStreamReader(CreateStream("123"))
        .Peek
        Actual = .Read
    End With
    
    Assert.That Actual, Iz.EqualTo(49)
End Sub

Public Sub Peek_WithEmptyStream_ReturnsMinusOne()
    Dim Actual As Long
    
    With Cor.NewStreamReader(CreateStream)
        Actual = .Peek
    End With
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub Peek_WhenAtEndOfStream_ReturnsMinusOne()
    Dim Actual As Long
    
    With Cor.NewStreamReader(CreateStream("123"))
        .ReadLine
        Actual = .Peek
    End With
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub ReadToEnd_WithLineBreaks_ReturnsAllData()
    Const Expected As String = "123" & vbCrLf & "456" & vbCrLf
    Dim Actual As String
    
    With Cor.NewStreamReader(CreateStream(Expected))
        Actual = .ReadToEnd
    End With
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadToEnd_WithLargeData_ReadsAllData()
    Dim Expected As String
    Expected = String$(20000, "1") & vbCrLf
    
    Dim Actual As String
    With Cor.NewStreamReader(CreateStream(Expected))
        Actual = .ReadToEnd
    End With
    
    Assert.That Len(Actual), Iz.EqualTo(20002)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadToEnd_WhenCalledAfterReadLine_ReadsRemainingData()
    Const Expected As String = "12345"
    
    Dim Actual As String
    With Cor.NewStreamReader(CreateStream("abcd" & vbCrLf & Expected))
        .ReadLine
        Actual = .ReadToEnd
    End With
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub NewStreamReader_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    
    NewStreamReader Nothing
    
    AssertArgumentNullException Err, "Source"
End Sub

Public Sub NewStreamReader_WithBufferSizeOfZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    NewStreamReader CreateStream("123"), BufferSize:=0
    
    AssertArgumentOutOfRangeException Err, "BufferSize"
End Sub

Public Sub NewStreamReader_WhenWithoutEncoding_DefaultsToUTF8Encoding()
    Dim Actual As Encoding
    
    With Cor.NewStreamReader(CreateStream)
        Set Actual = .CurrentEncoding
    End With
    
    Assert.That TypeOf Actual Is UTF8Encoding, Iz.True
End Sub

Public Sub ReadLine_WithGivenEncoding_ReturnsDecodedData()
    Const Expected As String = "a?b?c"
    
    Dim Actual As String
    With NewStreamReader(CreateStream("a" & Chr$(128) & "b" & Chr$(129) & "c"), Encoding.ASCII)
        Actual = .ReadLine
    End With
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CloseReader_WithLeaveStreamOpen_DoesNotCloseStream()
    Dim Stream As New FakeStream
    
    With NewStreamReader(Stream, LeaveOpen:=True)
        .CloseReader
    End With
    
    Assert.That Stream.IsClosed, Iz.False
End Sub

Public Sub EndOfStream_WithEmptyStream_ReturnsTrue()
    Dim Actual As Boolean
    
    With Cor.NewStreamReader(CreateStream)
        Actual = .EndOfStream
    End With
    
    Assert.That Actual, Iz.True
End Sub

Public Sub EndOfStream_WithUnreadData_ReturnsFalse()
    Dim Actual As Boolean
    
    With Cor.NewStreamReader(CreateStream("1245"))
        Actual = .EndOfStream
    End With
    
    Assert.That Actual, Iz.False
End Sub

Public Sub EndOfStream_WithAllDataRead_ReturnsTrue()
    Dim Actual As Boolean
    
    With Cor.NewStreamReader(CreateStream("12345"))
        .ReadToEnd
        Actual = .EndOfStream
    End With
    
    Assert.That Actual, Iz.True
End Sub

Private Function CreateStream(Optional ByVal Data As String) As Stream
    Dim Stream As New MemoryStream
    
    If Len(Data) > 0 Then
        Dim Bytes() As Byte
        Bytes = StrConv(Data, vbFromUnicode)
        Stream.WriteBlock Bytes, 0, UBound(Bytes) + 1
        Stream.Position = 0
    End If
    
    Set CreateStream = Stream
End Function

