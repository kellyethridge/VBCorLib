VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBParseHexTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBParseHexTests
'
Option Explicit

Private b As BigInteger


Public Sub Setup()
    Set b = Nothing
End Sub


Public Sub TryEmpty()
    Assert.IsFalse BigInteger.TryParse("&h", b)
End Sub

Public Sub TryZero()
    Assert.IsTrue BigInteger.TryParse("&h0", b)
    AssertNumber 0, 0, Zero, NewBytes()
End Sub

Public Sub TryNegativeZero()
    Assert.IsTrue BigInteger.TryParse("&h00", b)
    AssertNumber 0, 0, Zero, NewBytes()
End Sub

Public Sub TryOne()
    Assert.IsTrue BigInteger.TryParse("0x1", b)
    AssertNumber 1, 2, Positive, NewBytes(1)
End Sub

Public Sub TryFF()
    Assert.IsTrue BigInteger.TryParse("0xff", b)
    AssertNumber 1, 2, Positive, NewBytes(&HFF)
End Sub

Public Sub TryFFFF()
    Assert.IsTrue BigInteger.TryParse("0xffff", b)
    AssertNumber 2, 2, Negative, NewBytes(&HFF, &HFF)
End Sub

Public Sub Try0FFFF()
    Assert.IsTrue BigInteger.TryParse("0x0ffff", b)
    AssertNumber 2, 3, Negative, NewBytes(&HFF, &HFF)
End Sub

Public Sub Try0000FFFF()
    Assert.IsTrue BigInteger.TryParse("0x0000ffff", b)
    AssertNumber 2, 0, Positive, NewBytes(&HFF, &HFF)
End Sub

Public Sub Try0000ffffffffffffffff()
    Assert.IsTrue BigInteger.TryParse("0x0000ffffffffffffffff", b)
    AssertNumber 8, 0, Positive, NewBytes(&HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF)
End Sub

Public Sub Try80()
    Assert.IsTrue BigInteger.TryParse("&h80", b)
    AssertNumber 1, 2, Positive, NewBytes(&H80)
End Sub

Public Sub TryFF80()
    Assert.IsTrue BigInteger.TryParse("&hff80", b)
    AssertNumber 2, 2, Negative, NewBytes(&H80, &HFF)
End Sub

Public Sub Try8000()
    Assert.IsTrue BigInteger.TryParse("&h8000", b)
    AssertNumber 2, 2, Negative, NewBytes(0, &H80)
End Sub

Public Sub Try00008000()
    Assert.IsTrue BigInteger.TryParse("&h00008000", b)
    AssertNumber 2, 3, Positive, NewBytes(0, &H80)
End Sub

Public Sub Try0000ff()
    Assert.IsTrue BigInteger.TryParse("0x0000ff", b)
    AssertNumber 1, 2, Positive, NewBytes(&HFF)
End Sub

Public Sub TryFF7FFFFF()
    Assert.IsTrue BigInteger.TryParse("0xff7fffff", b)
    AssertNumber 4, 4, Negative, NewBytes(&HFF, &HFF, &H7F, &HFF)
End Sub

Public Sub Try00000000()
    Assert.IsTrue BigInteger.TryParse("0x00000000", b)
    AssertNumber 0, 0, Zero, NewBytes()
End Sub

Public Sub TryInvalid()
    Assert.IsFalse BigInteger.TryParse("0xaa.a", b)
    Assert.IsNothing b
End Sub

Public Sub Try80000000()
    Assert.IsTrue BigInteger.TryParse("&h80000000", b)
    AssertNumber 4, 0, Negative, NewBytes(0, 0, 0, &H80)
End Sub

Public Sub Try800000000000()
    Assert.IsTrue BigInteger.TryParse("&h800000000000", b)
    AssertNumber 6, 5, Negative, NewBytes(0, 0, 0, 0, 0, &H80)
End Sub

Public Sub TryMinus7FFF()
    Assert.IsTrue BigInteger.TryParse("-&h7fff", b)
    AssertNumber 2, 0, Negative, NewBytes(&H1, &H80)
End Sub

Public Sub TryMinus80000000()
    Assert.IsTrue BigInteger.TryParse("-&h80000000", b)
    AssertNumber 4, 0, Negative, NewBytes(0, 0, 0, &H80)
End Sub




Private Sub AssertNumber(ByVal Precision As Long, ByVal Length As Long, ByVal Sign As Sign, ByRef Bytes() As Byte)
    Assert.AreEqual Precision, b.Precision, "Wrong Precision"
'    Assert.AreEqual Length, b.Length, "Wrong Length"
    Assert.AreEqual Sign, b.Sign, "Wrong Sign"
    Assert.AreEqual Bytes, b.ToArray, "Wrong Array"
End Sub




