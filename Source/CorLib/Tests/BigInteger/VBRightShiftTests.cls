VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBRightShiftTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBShiftRightTests
'
Option Explicit

Private b As BigInteger

Public Sub Setup()
    Set b = Nothing
End Sub

Public Sub Shift0()
    Set b = BigInteger.Zero.ShiftRight(0)
    Assert.IsNotNothing b
    Assert.IsTrue b.Equals(BigInteger.Zero)
    
    Set b = BigInteger.Ten.ShiftRight(0)
    Assert.IsTrue b.Equals(BigInteger.Ten)
    
    Set b = Cor.BInt("0x0000ffffffff00000000").ShiftRight(0)
    Assert.AreEqual "ffffffff00000000", b.ToString("x")
End Sub

Public Sub Shift8()
    Set b = Cor.NewBigInteger(&HFFFF00).ShiftRight(8)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "ffff", b.ToString("x")
    
    Set b = Cor.NewBigInteger(&H8000).ShiftRight(8)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "ff80", b.ToString("x")
    
    Set b = Cor.BInt("0x1fffffff0000").ShiftRight(8)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1fffffff00", b.ToString("x")
End Sub

Public Sub Shift1()
    Set b = Cor.NewBigInteger(&HFFFF00).ShiftRight(1)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "7fff80", b.ToString("x")
    
    Set b = Cor.NewBigInteger(&H8000).ShiftRight(1)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "c000", b.ToString("x")
    
    Set b = Cor.NewBigInteger(&H10000).ShiftRight(1)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "8000", b.ToString("x")
    
    Set b = Cor.NewBigInteger(&HFFF10000).ShiftRight(1)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "fff88000", b.ToString("x")
    
    Set b = Cor.NewBigInteger(&HFFFF0000).ShiftRight(1)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "8000", b.ToString("x")
End Sub

Public Sub Shift15()
    Set b = Cor.NewBigInteger(&HF000000).ShiftRight(15)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1e00", b.ToString("x")
    
    Set b = Cor.BInt("0xffff0f000000").ShiftRight(15)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "fffe1e00", b.ToString("x")
End Sub

Public Sub Shift16()
    Set b = Cor.BInt("0xffff0f000000").ShiftRight(16)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "ffff0f00", b.ToString("x")
End Sub

Public Sub ShiftTooFar()
    Set b = Cor.BInt("0xffffffffffffffffffffffff").ShiftRight(96)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-1", b.ToString
    
    Set b = Cor.BInt("0x0000ffffffffffffffffffffffff").ShiftRight(96)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub Shift1EqualsDivideBy2()
    Dim b1 As BigInteger
    Dim b2 As BigInteger
    Set b1 = Cor.BInt(&H94CC&)
    Set b2 = b1
    
    Set b1 = b1.Divide(BInt(2))
    Set b2 = b2.ShiftRight(1)
    
    Assert.AreEqual b1.Sign, b2.Sign
    Assert.AreEqual b1.ToString("X"), b2.ToString("X")
End Sub

Public Sub Shift11()
    Dim b As BigInteger
    Set b = BInt("0x0000AB54A98CEB1F0800").ShiftRight(11)
    Assert.AreEqual "156A95319D63E1", b.ToString("X")
End Sub

