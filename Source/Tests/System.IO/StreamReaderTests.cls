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
' StreamReaderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub Read_WhenCalledWithEmptyStream_ReturnsMinusOne()
    Dim Reader As StreamReader
    Set Reader = MakeEmptyReader
    
    Dim Actual As Long
    Actual = Reader.Read
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub Read_WhenCalledWithData_ReturnsCharactersInOrder()
    Dim Reader As StreamReader
    Set Reader = MakeReader("123")
    
    Assert.That Reader.Read, Iz.EqualTo(49)
    Assert.That Reader.Read, Iz.EqualTo(50)
    Assert.That Reader.Read, Iz.EqualTo(51)
    Assert.That Reader.Read, Iz.EqualTo(-1)
End Sub

Public Sub ReadLine_WhenCalledWithEmptyStream_ReturnsNullString()
    Dim Reader As StreamReader
    Set Reader = MakeEmptyReader
    
    Dim Actual As String
    Actual = Reader.ReadLine
    
    Assert.That CorString.IsNull(Actual), Iz.True
End Sub

Public Sub ReadLine_WhenCalledWithEmptyLine_ReturnsEmptyLine()
    Dim Reader As StreamReader
    Set Reader = MakeReader(vbCrLf)
    
    Dim Actual As String
    Actual = Reader.ReadLine
    
    Assert.That CorString.IsNull(Actual), Iz.False
    Assert.That Actual, Iz.Empty
End Sub

Public Sub ReadLine_WhenNoMoreData_ReturnsNullString()
    Dim Reader As StreamReader
    Set Reader = MakeReader("123")
    
    Dim Actual As String
    Reader.ReadLine
    Actual = Reader.ReadLine
    
    Assert.That CorString.IsNull(Actual), Iz.True
End Sub

Public Sub ReadLine_WhenCalledWithMultipleLinesOfData_ReturnsLinesOfData()
    Dim Reader As StreamReader
    Set Reader = MakeReader("1234" & vbCrLf & "6789")
        
    Assert.That Reader.ReadLine, Iz.EqualTo("1234")
    Assert.That Reader.ReadLine, Iz.EqualTo("6789")
    Assert.That CorString.IsNull(Reader.ReadLine), Iz.True
End Sub

Public Sub ReadLine_WithLargeData_ReadsData()
    Dim Reader As StreamReader
    Dim Expected As String
    Expected = Space$(10000)
    Set Reader = MakeReader(Expected)
    
    Dim Actual As String
    Actual = Reader.ReadLine
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadBlock_WithCountOfZero_ReturnsNoData()
    Dim Reader As StreamReader
    Dim Chars(100) As Integer
    Set Reader = MakeReader("1234")
    
    Dim Actual As Long
    Actual = Reader.ReadBlock(Chars, 0, 0)
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub ReadBlock_WithNullArray_ThrowsArgumentNullException()
    Dim Reader As StreamReader
    Set Reader = MakeReader("1234")
    
    On Error Resume Next
    Reader.ReadBlock NullChars, 0, 1
    
    AssertArgumentNullException Err, "Buffer"
End Sub

Public Sub ReadBlock_WithIndexLessThanBufferLBound_ThrowsArgumentOutOfRangeException()
    Dim Reader As StreamReader
    Dim Chars(1 To 10) As Integer
    Set Reader = MakeReader("1234")
    
    On Error Resume Next
    Reader.ReadBlock Chars, 0, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub ReadBlock_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Reader As StreamReader
    Dim Chars(0 To 10) As Integer
    Set Reader = MakeReader("1234")
    
    On Error Resume Next
    Reader.ReadBlock Chars, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub ReadBlock_ThroughBufferSize_ReadsData()
    Dim Reader As StreamReader
    Dim Expected As String
    Dim Chars(0 To 127) As Integer
    Dim Output As New StringBuilder
    Dim CharsRead As Long
    Expected = String$(600, 48)
    Set Reader = Cor.NewStreamReader(MakeStream(Expected), BufferSize:=192)
    
    Do
        CharsRead = Reader.ReadBlock(Chars, 0, 128)
        Output.Append Chars, 0, CharsRead
    Loop While CharsRead > 0
    
    Assert.That Output.ToString, Iz.EqualTo(Expected)
End Sub

Public Sub Peek_WhenCalled_ReturnsCurrentCharacter()
    Dim Reader As StreamReader
    Set Reader = MakeReader("123")
    
    Dim Actual As Long
    Actual = Reader.Peek
    
    Assert.That Actual, Iz.EqualTo(Asc("1"))
End Sub

Public Sub Peek_WhenCalled_DoesNotAdvanceStreamPosition()
    Dim Reader As StreamReader
    Set Reader = MakeReader("123")
    
    Dim Actual As Long
    Reader.Peek
    Actual = Reader.Read
    
    Assert.That Actual, Iz.EqualTo(Asc("1"))
End Sub

Public Sub Peek_WithNoData_ReturnsMinusOne()
    Dim Reader As StreamReader
    Set Reader = MakeEmptyReader
    
    Dim Actual As Long
    Actual = Reader.Peek
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub Peek_WhenAtEndOfStream_ReturnsMinusOne()
    Dim Reader As StreamReader
    Set Reader = MakeReader("123")
    
    Dim Actual As Long
    Reader.ReadLine
    Actual = Reader.Peek
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub ReadToEnd_WithLineBreaks_ReturnsAllData()
    Const Expected As String = "123" & vbCrLf & "456" & vbCrLf
    Dim Reader As StreamReader
    Set Reader = MakeReader(Expected)
    
    Dim Actual As String
    Actual = Reader.ReadToEnd
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadToEnd_WithLargeData_ReadsAllData()
    Dim Expected As String
    Dim Reader As StreamReader
    Expected = String$(20000, "1") & vbCrLf
    Set Reader = MakeReader(Expected)
    
    Dim Actual As String
    Actual = Reader.ReadToEnd
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadToEnd_WhenCalledAfterReadLine_ReadsRemainingData()
    Const Expected As String = "12345"
    Dim Reader As StreamReader
    Set Reader = MakeReader("abcd" & vbCrLf & Expected)
    
    Dim Actual As String
    Reader.ReadLine
    Actual = Reader.ReadToEnd
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub NewStreamReader_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    
    NewStreamReader Nothing
    
    AssertArgumentNullException Err, "Source"
End Sub

Public Sub NewStreamReader_WithBufferSizeOfZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    NewStreamReader MakeStream("123"), BufferSize:=0
    
    AssertArgumentOutOfRangeException Err, "BufferSize"
End Sub

Public Sub NewStreamReader_WhenWithoutEncoding_DefaultsToUTF8Encoding()
    Dim Reader As StreamReader
    Set Reader = MakeEmptyReader
    
    Dim Actual As Encoding
    Set Actual = Reader.CurrentEncoding
    
    Assert.That TypeOf Actual Is UTF8Encoding, Iz.True
End Sub

Public Sub ReadLine_WithGivenEncoding_ReturnsDecodedData()
    Const Expected As String = "a?b?c"
    Dim Reader As StreamReader
    Set Reader = Cor.NewStreamReader(MakeStream(t("a\u0080b\u0081c")), Encoding.ASCII)
    
    Dim Actual As String
    Actual = Reader.ReadLine
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CloseReader_WithLeaveStreamOpen_DoesNotCloseStream()
    Dim Reader As StreamReader
    Dim Stream As New FakeStream
    Set Reader = Cor.NewStreamReader(Stream, LeaveOpen:=True)
    
    Reader.CloseReader
    
    Assert.That Stream.IsClosed, Iz.False
End Sub

Public Sub EndOfStream_WithEmptyStream_ReturnsTrue()
    Dim Reader As StreamReader
    Set Reader = MakeEmptyReader
    
    Dim Actual As Boolean
    Actual = Reader.EndOfStream
    
    Assert.That Actual, Iz.True
End Sub

Public Sub EndOfStream_WithUnreadData_ReturnsFalse()
    Dim Reader As StreamReader
    Set Reader = MakeReader("12345")
    
    Dim Actual As Boolean
    
    Actual = Reader.EndOfStream
    
    Assert.That Actual, Iz.False
End Sub

Public Sub EndOfStream_WithAllDataRead_ReturnsTrue()
    Dim Reader As StreamReader
    Set Reader = MakeReader("12345")
    
    Dim Actual As Boolean
    Reader.ReadToEnd
    Actual = Reader.EndOfStream
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Read_WithLittleEndianUnicodeEncodingBOM_SetsCorrectEncoding()
    Dim Reader As StreamReader
    Set Reader = MakeReader(t("\u00ff\u00fedata"))
    
    Dim Actual As Encoding
    Reader.Read
    Set Actual = Reader.CurrentEncoding
    
    Assert.That Actual, Equals(Cor.NewUnicodeEncoding(False, True))
End Sub

Public Sub Read_WithBigEndianUnicodeEncodingBOM_SetsCorrectEncoding()
    Dim Reader As StreamReader
    Set Reader = MakeReader(t("\u00fe\u00ffdata"))
    
    Dim Actual As Encoding
    Reader.Read
    Set Actual = Reader.CurrentEncoding
    
    Assert.That Actual, Equals(Cor.NewUnicodeEncoding(True, True))
End Sub

Public Sub Read_WithUTF8EncodingBOM_SetsCorrectEncoding()
    Dim Reader As StreamReader
    Set Reader = MakeReader(t("\u00ef\u00bb\u00bf"))
    
    Dim Actual As Encoding
    Reader.Read
    Set Actual = Reader.CurrentEncoding
    
    Assert.That Actual, Equals(Cor.NewUTF8Encoding(True))
End Sub

Public Sub Read_WithPartialUTF8EncodingBOM_DoesNotAlterEncoding()
    Dim Reader As StreamReader
    Set Reader = MakeReader(t("\u00ef\u00bbz"))
    
    Dim Actual As Encoding
    Reader.Read
    Set Actual = Reader.CurrentEncoding
    
    Assert.That Actual, Equals(Cor.NewUTF8Encoding(True))
End Sub

Public Sub Read_WithPartialUTF8EncodingBOM_ReturnsFirstCharacterAsInvalid()
    Dim Reader As StreamReader
    Set Reader = MakeReader(t("\u00ef\u00bb0"))
    
    Dim Actual As Long
    Actual = Reader.Read
    
    Assert.That Actual, Iz.EqualTo(&HFFFD&)
End Sub

Public Sub Read_WithMultipleLittleEndianUnicodeEncodingBOM_ReadsSecondBOMAsCharacter()
    Dim Reader As StreamReader
    Set Reader = MakeReader(t("\u00ff\u00fe\u00ff\u00fe"))
    
    Dim Actual As Long
    Actual = Reader.Read
    
    Assert.That Actual, Iz.EqualTo(&HFEFF&)
End Sub

Public Sub Read_WithPartialUTF8EncodingBOM_ReturnsMinusOne()
    Dim Reader As StreamReader
    Set Reader = MakeReader(t("\u00ef\u00bb"))
    
    Dim Actual As Long
    Actual = Reader.Read
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub Read_WithLittleEndianUTF32BOM_SetsUTF32Encoding()
    Dim Reader As StreamReader
    Dim Actual As Encoding
    Dim Ch      As Long
    
    Set Reader = MakeReader(t("\u00ff\u00fe\u0000\u0000a\u0000\u0000\u0000"))
    
    Ch = Reader.Read
    Set Actual = Reader.CurrentEncoding
    
    Assert.That Actual, Equals(Cor.NewUTF32Encoding(False, True))
    Assert.That Ch, Iz.EqualTo(97)
End Sub

Public Sub Read_WithBigEndianUTF32BOM_SetsUTF32Encoding()
    Dim Reader As StreamReader
    Dim Actual As Encoding
    Dim Ch      As Long
    
    Set Reader = MakeReader(t("\u0000\u0000\u00fe\u00ff\u0000\u0000\u0000a"))
    
    Ch = Reader.Read
    Set Actual = Reader.CurrentEncoding
    
    Assert.That Actual, Equals(Cor.NewUTF32Encoding(True, True))
    Assert.That Ch, Iz.EqualTo(97)
End Sub


Private Function MakeEmptyReader() As StreamReader
    Set MakeEmptyReader = Cor.NewStreamReader(MakeEmptyStream)
End Function

Private Function MakeReader(ByRef Data As String) As StreamReader
    Set MakeReader = Cor.NewStreamReader(MakeStream(Data))
End Function

Private Function MakeEmptyStream() As Stream
    Set MakeEmptyStream = New MemoryStream
End Function

Private Function MakeStream(Optional ByRef Data As String) As Stream
    Dim Stream As New MemoryStream
    
    If Len(Data) > 0 Then
        Dim Bytes() As Byte
        Bytes = StrConv(Data, vbFromUnicode)
        Stream.WriteBlock Bytes, 0, UBound(Bytes) + 1
        Stream.Position = 0
    End If
    
    Set MakeStream = Stream
End Function


