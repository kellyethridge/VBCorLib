VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileStreamTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' FileStreamTests
'
Option Explicit

Private mBuilder As IOHelper

Public Sub Setup()
    Set mBuilder = New IOHelper
End Sub

Public Sub Teardown()
    Set mBuilder = Nothing
End Sub

Public Sub NewFileStream_WithEmptyPath_ThrowsArgumentException()
    On Error Resume Next
    
    NewFileStream "", Create
    
    AssertArgumentException Err, "Path"
End Sub

Public Sub CanRead_WithFileModeAppend_ReturnsFalse()
    Dim FileName As String
    FileName = mBuilder.CreateFile("Test.txt", "Hello")
    
    With NewFileStream(FileName, FileMode.Append)
        Assert.That .CanRead, Iz.False
        Assert.That .CanWrite, Iz.True
        Assert.That .CanSeek, Iz.True
    End With
End Sub

Public Sub CanRead_WithFileModeCreate_ReturnsTrue()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "Test.txt")
    
    With NewFileStream(FileName, FileMode.Create)
        Assert.That .CanRead, Iz.True
        Assert.That .CanWrite, Iz.True
        Assert.That .CanSeek, Iz.True
    End With
End Sub

Public Sub CanWrite_WithFileModeCreate_ReturnsTrue()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "Test.txt")
    
    Dim Actual As Boolean
    With NewFileStream(FileName, FileMode.Create)
        Actual = .CanWrite
    End With
    
    Assert.That Actual, Iz.True
End Sub

Public Sub NewFileStream_WithBufferSizeLessThanOne_ThrowsArgumentOutOfRangeException()
    Dim FileName As String
    FileName = Path.Combine(mBuilder.TempFolder, "Test.txt")
    
    On Error Resume Next
    With NewFileStream(FileName, Create, buffersize:=0)
    End With
    
    AssertArgumentOutOfRangeException Err, "BufferSize"
End Sub

Public Sub ReadByte_WhenCalled_ReturnsBytes()
    Dim FileName As String
    FileName = mBuilder.CreateFile(Data:="12")
    
    With NewFileStream(FileName, FileMode.OpenExisting)
        Assert.That .ReadByte, Iz.EqualTo(Asc("1"))
        Assert.That .ReadByte, Iz.EqualTo(Asc("2"))
        Assert.That .ReadByte, Iz.EqualTo(-1)
    End With
End Sub

'Private fs As FileStream
'
'Public Sub Setup()
'    Set fs = NewFileStream(App.Path & "\hello.txt", FileMode.OpenExisting)
'End Sub
'
'Public Sub Teardown()
'    fs.CloseStream
'End Sub
'
'Public Sub TestLength()
'    Assert.AreEqual 5@, fs.Length
'    fs.CloseStream
'End Sub
'
'Public Sub TestReadByte()
'    Dim b As Byte
'
'    b = fs.ReadByte
'    Assert.AreEqual Asc("h"), b
'    Assert.AreEqual 1@, fs.Position
'End Sub
'
'Public Sub TestReadBytePassedEnd()
'    Dim i As Long
'
'    For i = 1 To 5
'        fs.ReadByte
'    Next i
'    Assert.AreEqual 5@, fs.Position
'    Assert.AreEqual -1, fs.ReadByte
'    Assert.AreEqual 5@, fs.Position
'End Sub
'
'Public Sub TestReadBlockPassedEnd()
'    Dim b(0) As Byte
'
'    fs.SeekPosition 0, FromEnd
'    Assert.AreEqual 0, fs.ReadBlock(b, 0, 1)
'End Sub
'
'Public Sub TestRead()
'    Dim b() As Byte
'    Dim i As Long
'
'    ReDim b(400)
'    Assert.AreEqual 0@, fs.Position
'    Assert.AreEqual 5, fs.ReadBlock(b, 0, 50)
'    For i = 0 To 4
'        Assert.AreEqual Choose(i + 1, 104, 101, 108, 108, 111), b(i)
'    Next i
'    Assert.AreEqual 5@, fs.Position
'End Sub
'
'Public Sub TestSeekBeforeFile()
'    On Error GoTo errTrap
'    fs.SeekPosition -1, FromBeginning
'    On Error GoTo 0
'    Assert.Fail "An exception should be thrown."
'
'errTrap:
'    Dim Ex As Exception
'    Assert.IsTrue Catch(Ex), "An exception should be caught."
'    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
'End Sub
'
'Public Sub TestReadOneByteAtATime()
'    Dim b As Long
'    Dim s As String
'
'    b = fs.ReadByte
'    Do While b >= 0
'        s = s & Chr$(b)
'        b = fs.ReadByte
'    Loop
'
'    Assert.AreEqual "hello", s
'End Sub
'
'Public Sub TestWriteSeekReadRead()
'    Dim f As String
'    Dim i As Long
'
'    fs.CloseStream
'    f = App.Path & "\WriteSeekReadRead.txt"
'
'    Set fs = NewFileStream(f, FileMode.OpenOrCreate)
'    fs.SeekPosition 0, FromBeginning
'    For i = 48 To 57
'        fs.WriteByte i
'    Next i
'    fs.SeekPosition -5, FromCurrent
'    Assert.AreEqual 53, fs.ReadByte
'    fs.SeekPosition 2, FromCurrent
'    Assert.AreEqual 8@, fs.Position, "Position should be 8."
'    Assert.AreEqual 56, fs.ReadByte
'
'End Sub
'
'Public Sub TestFileNotFound()
'    Dim NotFound As FileStream
'
'    On Error GoTo errTrap
'    Set NotFound = Cor.NewFileStream(App.Path & "\Cannot exist.txt", OpenExisting)
'    On Error GoTo 0
'    Assert.Fail "An exception should be thrown."
'
'errTrap:
'    Dim Ex As Exception
'    Assert.IsTrue Catch(Ex), "An exception should be caught."
'    Assert.IsTrue TypeOf Ex Is FileNotFoundException, "Wrong exception type caught."
'
'    Dim fex As FileNotFoundException
'    Set fex = Ex
'    Assert.AreEqual App.Path & "\Cannot exist.txt", fex.FileName
'End Sub
'
'Public Sub TestWriteSeekReadWrite()
'    fs.CloseStream
'
'    Dim f As String
'    f = App.Path & "\WriteSeekReadWrite.txt"
'
'    Set fs = NewFileStream(f, Create)
'
'    Dim b() As Byte
'    b = NewBytes(65, 65, 65, 65, 65)
'
'    fs.WriteBlock b, 0, 5
'
'    fs.Position = 0
'
'    Assert.AreEqual 65, fs.ReadByte
'
'    fs.WriteByte 66
'
'    fs.CloseStream
'
'    Assert.AreEqual 5, FileLen(f)
'
'End Sub
