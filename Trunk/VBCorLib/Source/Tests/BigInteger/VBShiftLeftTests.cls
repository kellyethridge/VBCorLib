VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBShiftLeftTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBShiftLeftTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub Shift0()
    Dim b As BigInteger
    Set b = BigInteger.Zero.ShiftLeft(0)
    
    Assert.IsNotNothing b
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub Shift8()
    Dim b As BigInteger
    Set b = BigInteger.One.ShiftLeft(8)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "256", b.ToString
    
    Set b = BigInteger.MinusOne.ShiftLeft(8)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "ff00", b.ToString("x")
    
    Set b = Cor.BInt("0x8000").ShiftLeft(8)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "ff800000", b.ToString("x")
    
    Set b = Cor.NewBigInteger(-2).ShiftLeft(8)
    Assert.AreEqual "fe00", b.ToString("x")
    Assert.AreEqual -1, b.Sign
    
    Set b = Cor.NewBigInteger(&HFFFF&).ShiftLeft(8)
    Assert.AreEqual "ffff00", b.ToString("x")
    Assert.AreEqual 1, b.Sign
End Sub

Public Sub Shift16()
    Dim b As BigInteger
    Set b = BigInteger.Ten.ShiftLeft(16)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "a0000", b.ToString("x")
    
    Set b = Cor.NewBigInteger(&HFFFF).ShiftLeft(16)
    Assert.AreEqual "ffff0000", b.ToString("x")
    Assert.AreEqual -1, b.Sign
    
    Set b = Cor.NewBigInteger(&HFFFF&).ShiftLeft(16)
    Assert.AreEqual "ffff0000", b.ToString("x")
    Assert.AreEqual 1, b.Sign
    
    Set b = Cor.NewBigInteger(-2).ShiftLeft(16)
    Assert.AreEqual "fffe0000", b.ToString("x")
    Assert.AreEqual -1, b.Sign
    
    Set b = BInt("0x0000ffffffff").ShiftLeft(16)
    Assert.AreEqual "ffffffff0000", b.ToString("x")
    Assert.AreEqual 1, b.Sign
End Sub

Public Sub Shift1()
    Dim b As BigInteger
    Set b = BigInteger.One.ShiftLeft(1)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "2", b.ToString
    
    Set b = Cor.NewBigInteger(&H8000).ShiftLeft(1)
    Assert.AreEqual "ffff0000", b.ToString("x")
    Assert.AreEqual -1, b.Sign
    
    Set b = Cor.NewBigInteger(&HFF00).ShiftLeft(1)
    Assert.AreEqual "fe00", b.ToString("x")
    Assert.AreEqual -1, b.Sign
    
    Set b = Cor.NewBigInteger(&HFFFF&).ShiftLeft(1)
    Assert.AreEqual "1fffe", b.ToString("x")
    Assert.AreEqual 1, b.Sign
    
    Set b = Cor.NewBigInteger(&HFFFF).ShiftLeft(1)
    Assert.AreEqual "fffe", b.ToString("x")
    Assert.AreEqual -1, b.Sign
    
    Set b = Cor.NewBigInteger(-2).ShiftLeft(1)
    Assert.AreEqual "fffc", b.ToString("x")
    Assert.AreEqual -1, b.Sign
    
    Set b = BInt("0x0000ffffffff").ShiftLeft(1)
    Assert.AreEqual "1fffffffe", b.ToString("x")
    Assert.AreEqual 1, b.Sign
End Sub

Public Sub Shift15()
    Dim b As BigInteger
    Set b = BigInteger.One.ShiftLeft(15)
    Assert.AreEqual "8000", b.ToString("x")
    Assert.AreEqual 1, b.Sign
    
    Set b = Cor.NewBigInteger(&HFFFF).ShiftLeft(15)
    Assert.AreEqual "8000", b.ToString("x")
    Assert.AreEqual -1, b.Sign
    
    Set b = Cor.NewBigInteger(&HFFFF&).ShiftLeft(15)
    Assert.AreEqual "7fff8000", b.ToString("x")
    Assert.AreEqual 1, b.Sign
End Sub

Public Sub Shift25()
    Dim b As BigInteger
    Set b = BigInteger.Ten.ShiftLeft(25)
    Assert.AreEqual "14000000", b.ToString("x")
    Assert.AreEqual 1, b.Sign
End Sub

Public Sub Shift11()
    Dim b As BigInteger
    Set b = BInt("0x00156A95319D63E1").ShiftLeft(11)
    Assert.AreEqual "AB54A98CEB1F0800", b.ToString("X")
End Sub
