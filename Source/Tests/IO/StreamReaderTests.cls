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

'Private ms As MemoryStream
'Private sr As StreamReader
'
'Const Q As Long = 10000

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub Setup()
    Set mIO = New IOHelper
End Sub

Public Sub Teardown()
    Set mIO = Nothing
End Sub

'Public Sub Setup()
'    Dim i As Long
'    Dim b() As Byte
'
'
'    ReDim b(Q - 1)
'    Set ms = New MemoryStream
'    For i = 0 To UBound(b)
'        b(i) = (i Mod 10) + 48
'    Next i
'    ms.WriteBlock b, 0, Q
'    ms.Position = 0
'    Set sr = Cor.NewStreamReader(ms)
'End Sub
'
'

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
    Dim Line As String
    With Cor.NewStreamReader(CreateStream(vbCrLf))
        Line = .ReadLine
        Assert.That cString.IsNull(Line), Iz.False
        Assert.That Line, Iz.Empty
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
            Actual.AppendChars Chars, 0, CharsRead
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
        Actual = .ReadLine
    End With
    
    Assert.That Len(Actual), Iz.EqualTo(20002)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'Public Sub TestReadToEnd()
'    Dim s As String
'    Dim st As Stream
'    Dim sr As StreamReader
'    Dim sw As StreamWriter
'    Dim sb As New StringBuilder
'    Dim i As Long
'
'    Set st = New MemoryStream
'
'    For i = 0 To 10000
'        sb.Append ChrW$(i And &HFFF)
'    Next i
'    s = sb.ToString
'
'    Set sw = newstreamwriter(st)
'    sw.WriteValue s
'    sw.Flush
'
'    st.Position = 0
'    Set sr = NewStreamReader(st)
'
'    Assert.AreEqual s, sr.ReadToEnd
'End Sub
'
'Public Sub TestReadLineThenReadToEnd()
'    Dim st As New MemoryStream
'    Dim sw As StreamWriter
'    Dim sr As StreamReader
'    Dim s As String
'    Dim strW As New StringWriter
'
'    strW.WriteLine String$(1000, ChrW$(&H2030))
'    strW.WriteValue String$(9000, ChrW$(&HA0))
'
'    Set sw = newstreamwriter(ms)
'    sw.WriteValue strW.ToString
'    sw.Flush
'
'    ms.Position = 0
'    Set sr = NewStreamReader(ms)
'
'    Assert.AreEqual String$(1000, ChrW$(&H2030)), sr.ReadLine
'    Assert.AreEqual String$(9000, ChrW$(&HA0)), sr.ReadToEnd
'
'    sr.CloseReader
'    sw.CloseWriter
'End Sub
'
'Public Sub TestReadEmptyLines()
'    Dim ms As New MemoryStream
'    Dim i As Long
'    For i = 1 To 100
'        ms.WriteByte 13
'    Next i
'    ms.Position = 0
'
'    Dim sr As StreamReader
'    Set sr = Cor.NewStreamReader(ms)
'
'    i = 0
'    Dim s As String
'    Do
'        s = sr.ReadLine
'        If cString.IsNull(s) Then Exit Do
'        Assert.AreEqual "", s
'        i = i + 1
'    Loop
'
'    Assert.AreEqual 100, i
'End Sub

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

