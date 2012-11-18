VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCryptoStreamWriteBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCryptoStreamWrite
'
Option Explicit

Private cs As CryptoStream
Private ms As MemoryStream


Public Sub Setup()
    Set ms = New MemoryStream
    Set cs = Cor.NewCryptoStream(ms, New ToBase64Transform, WriteMode)
End Sub


Public Sub TestWriteOneBlock()
    Dim b() As Byte
    b = NewArray(vbByte, 1, 2, 3)
    cs.WriteBlock b, 0, 3
    Assert.AreEqual 4, ms.Length
    b = ms.ToArray
    
    Assert.AreEqual 65, b(0)
    Assert.AreEqual 81, b(1)
    Assert.AreEqual 73, b(2)
    Assert.AreEqual 68, b(3)
End Sub

Public Sub TestOneByteAtATime()
    Dim b() As Byte
    b = NewArray(vbByte, 1, 2, 3)
    cs.WriteBlock b, 0, 1
    Assert.AreEqual 0, ms.Length
    cs.WriteBlock b, 1, 1
    Assert.AreEqual 0, ms.Length
    cs.WriteBlock b, 2, 1
    Assert.AreEqual 4, ms.Length
    
    b = ms.ToArray
    
    Assert.AreEqual 65, b(0)
    Assert.AreEqual 81, b(1)
    Assert.AreEqual 73, b(2)
    Assert.AreEqual 68, b(3)
End Sub

Public Sub WriteOneBytePartialBuffer()
    cs.WriteByte 1
    cs.FlushFinalBlock
    Assert.AreEqual 4, ms.Length
    
    Dim b() As Byte
    b = ms.ToArray
    Assert.AreEqual 65, b(0)
    Assert.AreEqual 81, b(1)
    Assert.AreEqual 61, b(2)
    Assert.AreEqual 61, b(3)
End Sub

Public Sub WriteTwoBytePartialBuffer()
    cs.WriteByte 1
    cs.WriteByte 2
    cs.FlushFinalBlock
    Assert.AreEqual 4, ms.Length
    
    Dim b() As Byte
    b = ms.ToArray
    Assert.AreEqual 65, b(0)
    Assert.AreEqual 81, b(1)
    Assert.AreEqual 73, b(2)
    Assert.AreEqual 61, b(3)
End Sub

Public Sub WriteFourByteBuffer()
    Dim b() As Byte
    b = NewArray(vbByte, 1, 2, 3, 4)
    cs.WriteBlock b, 0, 4
    Assert.AreEqual 4, ms.Length
    b = ms.ToArray
    Assert.AreEqual 65, b(0)
    Assert.AreEqual 81, b(1)
    Assert.AreEqual 73, b(2)
    Assert.AreEqual 68, b(3)
    
    cs.FlushFinalBlock
    Assert.AreEqual 8, ms.Length
    b = ms.ToArray
    Assert.AreEqual 65, b(0)
    Assert.AreEqual 81, b(1)
    Assert.AreEqual 73, b(2)
    Assert.AreEqual 68, b(3)
    Assert.AreEqual 66, b(4)
    Assert.AreEqual 65, b(5)
    Assert.AreEqual 61, b(6)
    Assert.AreEqual 61, b(7)
End Sub

'4
'65
'81
'73
'68
'8
'65
'81
'73
'68
'66
'65
'61
'61

