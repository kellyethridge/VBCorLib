VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestBuffer
'
Option Explicit

Public Sub TestBlockCopy()
    Dim src(11) As Byte
    Dim dst(2) As Long
    Dim i As Long
    
    For i = 4 To 7
        src(i) = 1
    Next i
    Buffer.BlockCopy src, 0, dst, 0, 12
    Assert.AreEqual 0, dst(0)
    Assert.AreEqual &H1010101, dst(1)
    Assert.AreEqual 0, dst(2)
End Sub

Public Sub TestByteLength()
    Dim src(10) As Long
    Assert.AreEqual 44, Buffer.ByteLength(src), "ByteLength should return 44."
End Sub

Public Sub TestGetByte()
    Dim src(10) As Long
    src(5) = &HFF0000
    Assert.AreEqual 255, Buffer.GetByte(src, 22), "GetByte should return 255."
End Sub

Public Sub TestSetByte()
    Dim src(10) As Long
    Buffer.SetByte src, 22, 255
    Assert.AreEqual &HFF0000, src(5), "SetByte should set value to 255."
End Sub
