VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestBinaryWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestBinaryWriter
'
Option Explicit

Private bw As BinaryWriter
Private ms As MemoryStream

Public Sub Setup()
    Set ms = New MemoryStream
    Set bw = Cor.NewBinaryWriter(ms)
End Sub

Public Sub TestByte()
    bw.WriteValue CByte(129)
    Assert.AreEqual 1@, ms.Length
    ms.Position = 0
    Assert.AreEqual 129, ms.ReadByte
End Sub

Public Sub Test10Bytes()
    Dim i As Long
    
    For i = 1 To 10
        bw.WriteValue CByte(i)
    Next i
    Assert.AreEqual 10@, ms.Length
    ms.Position = 0
    For i = 1 To 10
        Assert.AreEqual i, ms.ReadByte
    Next i
End Sub

Public Sub TestLong()
    bw.WriteValue 1&
    Assert.AreEqual 4@, ms.Length
    ms.Position = 0
    Assert.AreEqual 1, ms.ReadByte
    Assert.AreEqual 0, ms.ReadByte
    Assert.AreEqual 0, ms.ReadByte
    Assert.AreEqual 0, ms.ReadByte
End Sub

Public Sub Test10Longs()
    Dim i As Long
    For i = 1 To 10
        bw.WriteValue i
    Next i
    Assert.AreEqual 40@, ms.Length
    ms.Position = 0
    For i = 0 To 39
        If (i Mod 4) = 0 Then
            Assert.AreEqual i \ 4 + 1, ms.ReadByte
        Else
            Assert.AreEqual 0, ms.ReadByte
        End If
    Next i
End Sub

Public Sub TestString()
    Dim s As String
    
    s = "hello"
    bw.WriteValue s
    Assert.AreEqual 6@, ms.Length
    ms.Position = 0
    Assert.AreEqual 5, ms.ReadByte
    Assert.AreEqual 104, ms.ReadByte
    Assert.AreEqual 101, ms.ReadByte
    Assert.AreEqual 108, ms.ReadByte
    Assert.AreEqual 108, ms.ReadByte
    Assert.AreEqual 111, ms.ReadByte
End Sub

Public Sub Test200CharString()
    Dim s As String
    s = String$(200, "a")
    bw.WriteValue s
    Assert.AreEqual 202@, ms.Length
    ms.Position = 0
    
    Dim i As Long
    Assert.AreEqual 200, ms.ReadByte
    Assert.AreEqual 1, ms.ReadByte
    For i = 1 To 200
        Assert.AreEqual 97, ms.ReadByte
    Next i
End Sub

Public Sub TestWriteEmptyString()
    bw.WriteValue ""
    Assert.AreEqual 1@, ms.Length
    ms.Position = 0
    
    Assert.AreEqual 0, ms.ReadByte
End Sub

