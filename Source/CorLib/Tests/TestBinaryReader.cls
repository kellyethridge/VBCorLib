VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestBinaryReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   TestBinaryReader
'
Option Explicit

Private ms As MemoryStream
Private br As BinaryReader

Public Sub Setup()
    Set ms = New MemoryStream
    Set br = Cor.NewBinaryReader(ms)
End Sub

Public Sub TestReadByte()
    ms.WriteByte 129
    ms.Position = 0
    
    Assert.AreEqual 129, br.ReadByte
End Sub

Public Sub TestRead10Bytes()
    Dim i As Long
    For i = 1 To 10
        ms.WriteByte i
    Next i
    ms.Position = 0
    
    For i = 1 To 10
        Assert.AreEqual i, br.ReadByte
    Next i
End Sub

Public Sub TestReadLong()
    ms.WriteByte &HFF
    ms.WriteByte &H0
    ms.WriteByte &H80
    ms.WriteByte &H80
    ms.Position = 0
    
    Assert.AreEqual &H808000FF, br.ReadLong
End Sub

Public Sub TestReadString()
    Dim bw As BinaryWriter
    Set bw = Cor.NewBinaryWriter(ms)
    bw.WriteValue "hello" & ChrW$(&H2030)
    ms.Position = 0
    
    Assert.AreEqual "hello" & ChrW$(&H2030), br.ReadString
End Sub

Public Sub TestReadIntegerStringDouble()
    Dim bw As BinaryWriter
    Set bw = Cor.NewBinaryWriter(ms)
    bw.WriteValue 67
    bw.WriteValue "This is a string!"
    bw.WriteValue 12345.6789
    ms.Position = 0
    
    Assert.AreEqual 67, br.ReadInteger
    Assert.AreEqual "This is a string!", br.ReadString
    Assert.AreEqual 12345.6789, br.ReadDouble
    
End Sub

Public Sub TestReadChar()
    Dim bw As BinaryWriter
    Set bw = Cor.NewBinaryWriter(ms)
    bw.WriteChar &H2030
    ms.Position = 0 ' move passed the 1 byte string length descriptor
    
    Assert.AreEqual &H2030, br.Read
    Assert.AreEqual -1, br.Read
End Sub

Public Sub TestPeekChar()
    ms.WriteByte 97
    ms.Position = 0
    Assert.AreEqual 97, br.PeekChar
    Assert.AreEqual 97, br.Read
    Assert.AreEqual -1, br.Read
End Sub

Public Sub TestReadBigString()
    Dim i As Long
    Dim s As String
    Dim bw As BinaryWriter
    
    For i = 1 To 1000
        s = s & ChrW$(i)
    Next i
    
    Set bw = Cor.NewBinaryWriter(ms)
    bw.WriteValue s
    
    ms.Position = 0
    
    Assert.AreEqual s, br.ReadString
End Sub

Public Sub TestEndOfStream()
    Dim ms As New MemoryStream
    ms.WriteByte 1
    ms.WriteByte 2
    ms.Position = 0
    
    Dim br As BinaryReader
    Set br = Cor.NewBinaryReader(ms)
    
    On Error GoTo errTrap
    Dim i As Long
    i = br.ReadLong
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is EndOfStreamException, "Wrong exception type caught."
End Sub
