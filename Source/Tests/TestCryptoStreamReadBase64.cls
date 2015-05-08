VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCryptoStreamReadBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCryptoStreamRead
'
Option Explicit
Implements ICategorizable

Private cs As CryptoStream
Private ms As MemoryStream
Private buf() As Byte
Private org() As Byte

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub


Public Sub Setup()
    org = Encoding.ASCII.GetBytes("Hello, world!")
    buf = Encoding.ASCII.GetBytes(Convert.ToBase64String(org))
    Set ms = Cor.NewMemoryStream(buf, 0, CorArray.GetLength(buf), False)
    Set cs = Cor.NewCryptoStream(ms, New FromBase64Transform, ReadMode)
End Sub

Public Sub Teardown()
    cs.CloseStream
End Sub

Public Sub ReadZeroBytes()
    Dim b(0) As Byte
    Assert.AreEqual 0, cs.ReadBlock(b, 0, 0)
End Sub

Public Sub ReadOneByte()
    Dim b(0) As Byte
    Assert.AreEqual 1, cs.ReadBlock(b, 0, 1)
    Assert.AreEqual Asc("H"), b(0)
End Sub

Public Sub ReadTwoBytes()
    Dim b(1) As Byte
    Assert.AreEqual 2, cs.ReadBlock(b, 0, 2)
    Assert.AreEqual Asc("H"), b(0)
    Assert.AreEqual Asc("e"), b(1)
End Sub

Public Sub ReadTwoBytesOverConsecutiveReads()
    Dim b(0) As Byte
    Assert.AreEqual 1, cs.ReadBlock(b, 0, 1)
    Assert.AreEqual Asc("H"), b(0)
    Assert.AreEqual 1, cs.ReadBlock(b, 0, 1)
    Assert.AreEqual Asc("e"), b(0)
End Sub

Public Sub ReadWholeArrayBackOneByteAtATime()
    Dim b(0) As Byte
    
    Dim i As Long
    For i = 0 To UBound(org)
        Assert.AreEqual 1, cs.ReadBlock(b, 0, 1)
        Assert.AreEqual org(i), b(0)
    Next i
    
    Assert.AreEqual 0, cs.ReadBlock(b, 0, 1)
End Sub

Public Sub Read3BytesAtATime()
    Dim b(2) As Byte
    Dim i As Long
    For i = 0 To 3
        Assert.AreEqual 3, cs.ReadBlock(b, 0, 3)
        Dim j As Long
        For j = 0 To 2
            Assert.AreEqual org(i * 3 + j), b(j)
        Next j
    Next i
    
    Assert.AreEqual 1, cs.ReadBlock(b, 0, 3)
    Assert.AreEqual org(UBound(org)), b(0)
    Assert.AreEqual 0, cs.ReadBlock(b, 0, 1)
End Sub

Public Sub Read5BytesAtATime()
    Dim b(4) As Byte
    Dim i As Long
    For i = 0 To 1
        Assert.AreEqual 5, cs.ReadBlock(b, 0, 5)
        Dim j As Long
        For j = 0 To 4
            Assert.AreEqual org(i * 5 + j), b(j)
        Next j
    Next i
    
    Assert.AreEqual 3, cs.ReadBlock(b, 0, 3)
    Assert.AreEqual org(UBound(org) - 2), b(0)
    Assert.AreEqual org(UBound(org) - 1), b(1)
    Assert.AreEqual org(UBound(org)), b(2)
    Assert.AreEqual 0, cs.ReadBlock(b, 0, 1)
End Sub

Public Sub Read13BytesAtATime()
    Dim b(12) As Byte
    Assert.AreEqual 13, cs.ReadBlock(b, 0, 13)
    Assert.AreEqual org, b
    Assert.AreEqual 0, cs.ReadBlock(b, 0, 1)
    Assert.AreEqual -1, cs.ReadByte
End Sub

Public Sub ReadMoreThanLength()
    Dim b(100) As Byte
    Assert.AreEqual 13, cs.ReadBlock(b, 0, 100)
    Dim i As Long
    For i = 0 To 12
        Assert.AreEqual org(i), b(i)
    Next i
    For i = 13 To 100
        Assert.AreEqual 0, b(i)
    Next i
    Assert.AreEqual 0, cs.ReadBlock(b, 0, 100)
End Sub

Public Sub ReadAllUsingReadByte()
    Dim i As Long
    For i = 0 To UBound(org)
        Assert.AreEqual org(i), cs.ReadByte
    Next i
    Assert.AreEqual -1, cs.ReadByte
End Sub
