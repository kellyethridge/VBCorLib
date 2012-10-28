VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMemoryStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMemoryStream
'
Option Explicit
Implements AsyncCallback

Private ms As MemoryStream
Private mAsyncCallback As Boolean
Private b() As Byte

Public Sub Setup()
    Set ms = New MemoryStream
    mAsyncCallback = False
End Sub

Public Sub TestInitialPostion()
    Assert.AreEqual 0@, ms.Position
End Sub

Public Sub TestPositionAfterWriting1byte()
    ms.WriteByte 55
    Assert.AreEqual 1@, ms.Position
End Sub

Public Sub TestChangePosition()
    ms.Position = 1
    Assert.AreEqual 1@, ms.Position
    ms.Position = 0
    Assert.AreEqual 0@, ms.Position
End Sub

Public Sub TestReadPreviousByte()
    ms.WriteByte 55
    ms.Position = 0
    Assert.AreEqual 55, ms.ReadByte
End Sub

Public Sub TestReadByteAtEndOfStream()
    Assert.AreEqual -1, ms.ReadByte
End Sub

Public Sub TestWriteBlock()
    Dim i As Long
        
    ReDim b(9)
    For i = 0 To 9
        b(i) = i + 1
    Next i
    ms.WriteBlock b, 0, 10
    Assert.AreEqual 10@, ms.Position
End Sub

Public Sub TestReadBlock()
    Dim i As Long
    
    ReDim b(10)
    For i = 0 To 9
        b(i) = i + 1
    Next i
    b(10) = 13
    For i = 1 To 1000
        ms.WriteBlock b, 0, 11
    Next i
    ms.Position = 0
    
    Dim b2(10) As Byte
    Dim j As Long
    
    For j = 1 To 1000
        Assert.AreEqual ms.ReadBlock(b2, 0, 11), 11
        For i = 0 To 9
            Assert.AreEqual i + 1, b2(i)
        Next i
        Assert.AreEqual 13, b2(10)
    Next j
End Sub

Public Sub TestBeginWrite()
    Dim i As Long
    Dim ar As IAsyncResult
    
    ReDim b(9)
    For i = 0 To 9
        b(i) = i + 1
    Next i
    
    Set ar = ms.BeginWrite(b, 0, 10, Me)
    Assert.IsNotNothing ar, "ar should not be Nothing."
    Assert.IsTrue mAsyncCallback, "AsyncCallback should be executed."
    ms.EndWrite ar
    Assert.AreEqual 10@, ms.Length
End Sub

Public Sub TestBeginRead()
    Dim i As Long
    Dim ar As IAsyncResult
    
    ReDim b(9)
    For i = 0 To 9
        b(i) = i + 1
    Next i
    ms.WriteBlock b, 0, 10
    
    ReDim b(9)
    ms.SeekPosition 0, FromBeginning
    Set ar = ms.BeginRead(b, 0, 10, Me)
    Assert.IsNotNothing ar, "ar should not be Nothing."
    Assert.IsTrue mAsyncCallback, "AsyncCallback should be executed."
    Assert.AreEqual 10, ms.EndRead(ar)
    For i = 0 To 9
        Assert.AreEqual i + 1, b(i)
    Next i
End Sub

Public Sub TestReadToOneBytePassedEnd()
    Dim i As Long
    Dim b(10) As Byte
    
    For i = 1 To 10
        ms.WriteByte i
    Next i
    
    ms.Position = 8
    Assert.AreEqual 9, ms.ReadByte
    ms.Position = 8
    Assert.AreEqual 2, ms.ReadBlock(b, 0, 3)
    Assert.AreEqual 9, b(0)
    Assert.AreEqual 10, b(1)
End Sub



Private Sub AsyncCallback_Execute(ByVal ar As VBCorLib.IAsyncResult)
    mAsyncCallback = True
End Sub
