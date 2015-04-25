VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBMultiplyTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBMultiplyTests
'
Option Explicit
Implements ICategorizable

Private b As BigInteger

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Numerics"
End Sub

Public Sub MultiplyZeros()
    Set b = BigInteger.Multiply(Nothing, Nothing)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual 0, b.Precision
End Sub

Public Sub AnythingByZeroIsZero()
    Set b = BigInteger.Multiply(Nothing, BigInteger.One)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual 0, b.Precision
    
    Set b = BigInteger.Multiply(BigInteger.One, Nothing)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual 0, b.Precision
End Sub

Public Sub AnythingByOneIsItself()
    Set b = BigInteger.Multiply(BigInteger.Ten, BigInteger.One)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "10", b.ToString
    
    Set b = BigInteger.Multiply(BigInteger.One, BigInteger.Ten)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "10", b.ToString
End Sub

Public Sub AnythingByMinusOneIsNegated()
    Set b = BigInteger.Multiply(BigInteger.Ten, BigInteger.MinusOne)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-10", b.ToString
    
    Set b = BigInteger.Multiply(BigInteger.MinusOne, BigInteger.Ten)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-10", b.ToString
    
    Set b = BigInteger.Multiply(BigInteger.MinusOne, BigInteger.MinusOne)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub TwoTimesTwo()
    Set b = BigInteger.Multiply(Cor.NewBigInteger(2), Cor.NewBigInteger(2))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "4", b.ToString
End Sub

Public Sub TwoTimesTwoHundred()
    Set b = BigInteger.Multiply(Cor.NewBigInteger(2), Cor.NewBigInteger(200))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "400", b.ToString
End Sub

Public Sub MinusTwoTimesTwoHundred()
    Set b = BigInteger.Multiply(Cor.NewBigInteger(-2), Cor.NewBigInteger(200))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-400", b.ToString
End Sub

Public Sub TwoHundredTimesMinusTwo()
    Set b = BigInteger.Multiply(Cor.NewBigInteger(200), Cor.NewBigInteger(-2))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-400", b.ToString
End Sub

Public Sub MinusTwoTimesMinusTwoHundred()
    Set b = BigInteger.Multiply(Cor.NewBigInteger(-2), Cor.NewBigInteger(-200))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "400", b.ToString
End Sub

Public Sub LargePositives()
    Set b = BigInteger.Multiply(Cor.BInt("12345678901234567890"), Cor.BInt("987654321987654321"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "12193263124676116323609205901126352690", b.ToString
End Sub

Public Sub LargeNegatives()
    Set b = BigInteger.Multiply(Cor.BInt("-12345678901234567890"), Cor.BInt("-987654321987654321"))
    Assert.AreEqual "12193263124676116323609205901126352690", b.ToString
    Assert.AreEqual 1, b.Sign
End Sub

Public Sub LargeMixed()
    Set b = BigInteger.Multiply(Cor.BInt("-12345678901234567890"), Cor.BInt("987654321987654321"))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-12193263124676116323609205901126352690", b.ToString
End Sub

Public Sub LargePositives2()
    Set b = BigInteger.Multiply(Cor.BInt("951524275264729222930554919618623681"), Cor.BInt("951524275264729222930554919618623681"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "905398446418068188910700851949003793786466575990235544111568936693989761", b.ToString
End Sub


