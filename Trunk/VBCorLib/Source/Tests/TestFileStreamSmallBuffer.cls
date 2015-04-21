VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFileStreamSmallBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFileStreamSmallBuffer
'
Option Explicit

Private fs As FileStream

Public Sub Setup()
    Set fs = NewFileStream(App.Path & "\100bytes.txt", OpenExisting, buffersize:=8)
End Sub

Public Sub TearDown()
    fs.CloseStream
End Sub

Public Sub TestLength()
    Assert.AreEqual 100@, fs.Length, "Length should be 100 bytes."
End Sub

Public Sub TestSeekToChangeBuffer()
    Assert.AreEqual 0@, fs.Position
    Assert.AreEqual 49, fs.ReadByte
    Assert.AreEqual 11@, fs.SeekPosition(11, FromBeginning)
    Assert.AreEqual 50, fs.ReadByte
End Sub

Public Sub TestReadLarge()
    Dim b() As Byte
    Dim i As Long
    
    ReDim b(24)
    fs.Position = 12
    Assert.AreEqual 25, fs.ReadBlock(b, 0, 25)
    For i = 0 To 24
        Assert.AreEqual 48 + ((i + 3) Mod 10), b(i)
    Next i
    Assert.AreEqual 56, fs.ReadByte
End Sub

Public Sub TestReadAll()
    Dim b() As Byte
    Dim i As Long
    
    ReDim b(100)
    Assert.AreEqual 100, fs.ReadBlock(b, 0, 100)
    For i = 0 To 99
        Assert.AreEqual 48 + ((i + 1) Mod 10), b(i)
    Next i
    Assert.AreEqual 0, fs.ReadBlock(b, 0, 1)
    Assert.AreEqual -1, fs.ReadByte
End Sub

Public Sub ReadInconsistent()
    Dim b() As Byte
    Dim i As Long
    
    ReDim b(100)
    fs.ReadBlock b, 0, 6
    fs.ReadByte
    Assert.AreEqual 5, fs.ReadBlock(b, 0, 5)
    
    For i = 0 To 4
        Assert.AreEqual 48 + ((8 + i) Mod 10), b(i)
    Next i
End Sub

Public Sub TestReadOneByteAtATime()
    Dim i As Long
    Dim b As Long
    Dim j As Long
  
    
    
    b = fs.ReadByte
    Do While b >= 0
        j = j + 1
        i = (i + 1) Mod 10
        Assert.AreEqual 48 + i, b
        b = fs.ReadByte
    Loop
    
    Assert.AreEqual 100, j
End Sub

Public Sub TestReadBufferSizeAtATime()
    Dim b(7) As Byte
    Call fs.ReadBlock(b, 0, 8)
    Call fs.ReadBlock(b, 0, 1)
    Call fs.ReadBlock(b, 0, 1)
End Sub
