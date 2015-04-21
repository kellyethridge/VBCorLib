VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBToStringHexTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBToStringHexTests
'
Option Explicit

Private b As BigInteger

Public Sub Setup()
    Set b = Nothing
End Sub

Public Sub TestZero()
    Set b = New BigInteger
    Assert.AreEqual "0", b.ToString("x")
    Assert.AreEqual "0", b.ToString("X")
    Assert.AreEqual "0000", b.ToString("x4")
End Sub

Public Sub Test1To1000000000()
    Dim i As Long
    i = 1
    Do
        i = i * 10
        Set b = NewBigInteger(i)
        Assert.AreEqual LCase$(Hex$(i)), b.ToString("x")
        Assert.AreEqual Hex$(i), b.ToString("X")
    Loop While i < 1000000000
End Sub

Public Sub TestNegative8000()
    Set b = Cor.NewBigInteger(&HFFFF8000)
    Assert.AreEqual "8000", b.ToString("x")
    Assert.AreEqual "f8000", b.ToString("x5")
    Assert.AreEqual "ffff8000", b.ToString("x8")
    Assert.AreEqual "8000", b.ToString("x4")
End Sub

Public Sub TestNegativeffff7fff()
    Set b = Cor.NewBigInteger(&HFFFF7FFF)
    Assert.AreEqual "ffff7fff", b.ToString("x")
    Assert.AreEqualStrings "fffff7fff", b.ToString("x9"), vbBinaryCompare
    Assert.AreEqualStrings "FFFFF7FFF", b.ToString("X9"), vbBinaryCompare
End Sub

Public Sub Testffff8001()
    Set b = Cor.NewBigInteger(&HFFFF8001)
    Assert.AreEqual "8001", b.ToString("x")
    Assert.AreEqual "ffff8001", b.ToString("x8")
End Sub

Public Sub Test7FFF()
    Set b = Cor.NewBigInteger(&H7FFF)
    Assert.AreEqual "7fff", b.ToString("x4")
    Assert.AreEqual "07fff", b.ToString("x5")
    Assert.AreEqual "00007fff", b.ToString("x8")
End Sub

Public Sub TestFFFF0000()
    Set b = Cor.NewBigInteger(&HFFFF0000)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "ffff0000", b.ToString("x")
    'Assert.AreEqual "ffff0000", b.ToString("x8")
End Sub

