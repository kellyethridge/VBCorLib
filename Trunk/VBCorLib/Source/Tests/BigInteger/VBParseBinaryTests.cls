VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBParseBinaryTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBParseBinaryTests
'
Option Explicit

Private b As BigInteger


Public Sub Setup()
    Set b = Nothing
End Sub

Public Sub TestEmptyString()
    Assert.IsFalse BigInteger.TryParse("0b", b)
    Assert.IsNothing b
End Sub

Public Sub Test0()
    Assert.IsTrue BigInteger.TryParse("0b0", b)
    AssertNumber 0, 0, Zero, NewBytes()
End Sub

Public Sub Test1()
    Assert.IsTrue BigInteger.TryParse("0b1", b)
    AssertNumber 1, 2, Positive, NewBytes(1)
End Sub

Public Sub Test11()
    Assert.IsTrue BigInteger.TryParse("0b11", b)
    AssertNumber 1, 2, Positive, NewBytes(3)
End Sub

Public Sub Test1000()
    Assert.IsTrue BigInteger.TryParse("0b1000", b)
    AssertNumber 1, 2, Positive, NewBytes(&H8)
End Sub

Public Sub Test1111()
    Assert.IsTrue BigInteger.TryParse("0b1111", b)
    AssertNumber 1, 2, Positive, NewBytes(&HF)
End Sub

Public Sub Test111111111()
    Assert.IsTrue BigInteger.TryParse("0b111111111", b)
    AssertNumber 2, 3, Positive, NewBytes(&HFF, &H1)
End Sub

Public Sub Test8000()
    Assert.IsTrue BigInteger.TryParse("0b10000000000000000000000000000000", b)
    AssertNumber 4, 4, Negative, NewBytes(0, 0, 0, &H80)
End Sub

Public Sub Test008000()
    Assert.IsTrue BigInteger.TryParse("0b000000000000000010000000000000000000000000000000", b)
    AssertNumber 4, 5, Positive, NewBytes(0, 0, 0, &H80)
End Sub

Public Sub Test18000()
    Assert.IsTrue BigInteger.TryParse("0b110000000000000000000000000000000", b)
    AssertNumber 5, 5, Positive, NewBytes(0, 0, 0, &H80, 1)
End Sub

Public Sub Test10000()
    Assert.IsTrue BigInteger.TryParse("0b100000000000000000000000000000000", b)
    AssertNumber 5, 6, Positive, NewBytes(0, 0, 0, 0, 1)
End Sub

Public Sub TestInvalidFormat()
    Assert.IsFalse BigInteger.TryParse("0b12", b)
    Assert.IsNothing b
End Sub

Public Sub TestMinus111111100000000()
    Assert.IsTrue BigInteger.TryParse("-0b111111100000000", b)
    AssertNumber 2, 0, Negative, NewBytes(0, &H81)
End Sub



Private Sub AssertNumber(ByVal Precision As Long, ByVal Length As Long, ByVal Sign As Sign, ByRef Bytes() As Byte)
    Assert.AreEqual Precision, b.Precision, "Wrong Precision"
'    Assert.AreEqual Length, b.Length, "Wrong Length"
    Assert.AreEqual Sign, b.Sign, "Wrong Sign"
    Assert.AreEqual Bytes, b.ToArray, "Wrong Array"
End Sub



