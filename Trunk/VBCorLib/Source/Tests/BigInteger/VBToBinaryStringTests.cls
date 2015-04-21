VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBToBinaryStringTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBToBinaryStringTests
'
Option Explicit

Private b As BigInteger


Public Sub Setup()
    Set b = Nothing
End Sub

Public Sub TestEmpty()
    Set b = New BigInteger
    Assert.AreEqual "0", b.ToString("b")
    Assert.AreEqual "0000", b.ToString("b4")
End Sub

Public Sub Test0()
    Set b = NewBigInteger(0)
    Assert.AreEqual "0", b.ToString("b")
    Assert.AreEqual "00000000", b.ToString("b8")
End Sub

Public Sub Test1()
    Set b = NewBigInteger(1)
    Assert.AreEqual "1", b.ToString("b")
    Assert.AreEqual "01", b.ToString("b2")
    Assert.AreEqual "0001", b.ToString("b4")
End Sub

Public Sub Test11()
    Set b = NewBigInteger(3)
    Assert.AreEqual "11", b.ToString("b")
    Assert.AreEqual "000011", b.ToString("b6")
End Sub

Public Sub Test10000000()
    Set b = NewBigInteger(128)
    Assert.AreEqual "10000000", b.ToString("b")
End Sub

Public Sub Test1000000010000000()
    Set b = NewBigInteger(&H8080)
    Assert.AreEqual "1000000010000000", b.ToString("b")
    Assert.AreEqual "11111000000010000000", b.ToString("b20")
End Sub

Public Sub TestMinus1()
    Set b = NewBigInteger(-1)
    Assert.AreEqual "1111111111111111", b.ToString("b")
    Assert.AreEqual "11111111111111111111111111111111", b.ToString("b32")
End Sub

Public Sub Test7FFF()
    Set b = NewBigInteger(&H7FFF)
    Assert.AreEqual "111111111111111", b.ToString("B")
    Assert.AreEqual "00000000000000000111111111111111", b.ToString("b32")
End Sub

