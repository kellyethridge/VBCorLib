VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCryptoStreamFullBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCryptoStreamFullBase64
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Test1171BytesOneByteAtATime()
    Dim csWrt As CryptoStream
    Dim ms As New MemoryStream
    Set csWrt = Cor.NewCryptoStream(ms, New ToBase64Transform, WriteMode)
    
    Dim i As Long
    For i = 1 To 1171
        csWrt.WriteByte i And &HFF
    Next i
    Assert.AreEqual 1560, ms.Length
    csWrt.FlushFinalBlock
    Assert.AreEqual 1564, ms.Length
    
    Dim csRd As CryptoStream
    ms.Position = 0
    Set csRd = Cor.NewCryptoStream(ms, New FromBase64Transform, ReadMode)
    
    For i = 1 To 1171
        Assert.AreEqual i And &HFF, csRd.ReadByte
    Next i
    Assert.AreEqual -1, csRd.ReadByte
End Sub

Public Sub Test100ThirteenByteBlocks()
    Dim csWrt As CryptoStream
    Dim ms As New MemoryStream
    Set csWrt = Cor.NewCryptoStream(ms, New ToBase64Transform, WriteMode)
    
    Dim b(12) As Byte
    Dim r As New Random
    r.NextBytes b
    
    Dim i As Long
    For i = 1 To 100
        csWrt.WriteBlock b, 0, 13
    Next i
    Assert.AreEqual 1732, ms.Length
    csWrt.FlushFinalBlock
    Assert.AreEqual 1736, ms.Length
    
    Dim b2(12) As Byte
    Dim csRd As CryptoStream
    ms.Position = 0
    Set csRd = Cor.NewCryptoStream(ms, New FromBase64Transform, ReadMode)
    
    For i = 1 To 100
        Erase b2
        Assert.AreEqual 13, csRd.ReadBlock(b2, 0, 13)
        Dim j As Long
        For j = 0 To 12
            Assert.AreEqual b(j), b2(j), "i = " & i & "; j = " & j
        Next j
    Next i
    Assert.AreEqual 0, csRd.ReadBlock(b2, 0, 13)
End Sub

