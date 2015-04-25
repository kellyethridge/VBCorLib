VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBBitTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBBitTests
'
Option Explicit
Implements ICategorizable

Private b As BigInteger

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub Setup()
    Set b = Nothing
End Sub

Public Sub TestZero()
    Assert.IsFalse BigInteger.Zero.TestBit(0)
    Assert.IsFalse BigInteger.Zero.TestBit(1024)
End Sub

Public Sub TestOne()
    Set b = BigInteger.One
    Assert.IsTrue b.TestBit(0)
    Assert.IsFalse b.TestBit(1)
    Assert.IsFalse b.TestBit(8)
    Assert.IsFalse b.TestBit(1024)
End Sub

Public Sub TestMinusOne()
    Set b = BigInteger.MinusOne
    Assert.IsTrue b.TestBit(0)
    Assert.IsTrue b.TestBit(7)
    Assert.IsTrue b.TestBit(8)
    Assert.IsTrue b.TestBit(1024)
End Sub

Public Sub Test256()
    Set b = BInt(256)
    Assert.IsFalse b.TestBit(7)
    Assert.IsTrue b.TestBit(8)
    Assert.IsFalse b.TestBit(9)
End Sub

Public Sub SetBit0TrueOnZero()
    Set b = BigInteger.Zero.SetBit(0)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub SetBit8TrueOnZero()
    Set b = BigInteger.Zero.SetBit(8)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "256", b.ToString
End Sub

Public Sub SetBit256FalseOnZero()
    Set b = BigInteger.Zero.ClearBit(256)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub SetBit9TrueOnMinusOne()
    Set b = BigInteger.MinusOne.SetBit(9)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-1", b.ToString
End Sub

Public Sub SetBit90FalseOnFFFFFFFFFFFFFFFFFFFFFFFF()
    Set b = BigInteger.Parse("0xFFFFFFFFFFFFFFFFFFFFFFFF")
    Set b = b.ClearBit(90)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "FBFFFFFFFFFFFFFFFFFFFFFF", b.ToString("X")
End Sub

Public Sub SetBit90TrueOn0080()
    Set b = BigInteger.Parse("0x0080")
    Set b = b.SetBit(90)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "40000000000000000000080", b.ToString("X")
End Sub

Public Sub SetBit90TrueOn1()
    Set b = Cor.NewBigInteger(1)
    Set b = b.SetBit(90)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "40000000000000000000001", b.ToString("X")
End Sub

Public Sub ClearBit90On1()
    Set b = Cor.NewBigInteger(1)
    Set b = b.ClearBit(90)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub AndBits1()
    Set b = BigInteger.AndBits(BInt(&HFFFFFFFF), BInt(&H10101010))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "10101010", b.ToString("x")
End Sub

Public Sub AndBits2()
    Set b = BigInteger.AndBits(BInt("&HFFFFFFFFFFFF"), BInt(&HFFF))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "FFF", b.ToString("X")
End Sub

Public Sub AndBits3()
    Set b = BigInteger.AndBits(BInt("&H0000FFFFFFFFFFFF"), BInt("&HFFFFFF000000"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "FFFFFF000000", b.ToString("X")
End Sub

Public Sub AndBits4()
    Set b = BigInteger.AndBits(BInt(-1), BInt("&H80000000000000000000000000000000"))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "80000000000000000000000000000000", b.ToString("x")
End Sub

Public Sub AndBits5()
    Set b = BigInteger.AndBits(BInt(-1), Nothing)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub AndBits6()
    Set b = BigInteger.AndBits(BInt("&H0000FFFFFFFFFFFF"), BInt("&H0000FFFFFF000000"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "FFFFFF000000", b.ToString("X")
End Sub

Public Sub AndBits7()
    Set b = BigInteger.AndBits(BInt(-1), BInt("&H0000FFFFFF000000"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "FFFFFF000000", b.ToString("X")
End Sub

Public Sub OrBits1()
    Set b = BigInteger.OrBits(BInt(&HFFF0000), BInt(&HFFFF&))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "FFFFFFF", b.ToString("X")
End Sub

Public Sub OrBits2()
    Set b = BigInteger.OrBits(Nothing, BInt("0x000080000000000000000000000000000000000000000000"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "80000000000000000000000000000000000000000000", b.ToString("x")
End Sub

Public Sub OrBits3()
    Set b = BigInteger.OrBits(Nothing, BInt("0x80000000000000000000000000000000000000000000"))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "80000000000000000000000000000000000000000000", b.ToString("x")
End Sub

Public Sub NotBits1()
    Set b = BigInteger.NotBits(BInt(&HFFFE))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub NotBits2()
    Set b = BigInteger.NotBits(BInt(&HFFFE&))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub NotBit3()
    Set b = BigInteger.NotBits(BInt(1))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "fffe", b.ToString("x")
End Sub

