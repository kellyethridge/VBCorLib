VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBComparisonTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ComparisonTests
'
Option Explicit

Public Sub ZeroEqualsZero()
    Assert.AreEqual 0, BigInteger.Compare(Nothing, Nothing)
    Assert.That 1, Iz.Not.EqualTo(BigInteger.Compare(Nothing, Nothing))
    Assert.That -1, Iz.Not.EqualTo(BigInteger.Compare(Nothing, Nothing))
End Sub

Public Sub OneAndZero()
    Assert.AreEqual 1, BigInteger.Compare(BigInteger.One, BigInteger.Zero)
    Assert.AreEqual -1, BigInteger.Compare(BigInteger.Zero, BigInteger.One)
End Sub

Public Sub MinusOneAndOne()
    Assert.AreEqual 1, BigInteger.Compare(BigInteger.One, BigInteger.MinusOne)
    Assert.AreEqual -1, BigInteger.Compare(BigInteger.MinusOne, BigInteger.One)
End Sub

Public Sub MinusTenAndMinusOne()
    Assert.AreEqual 1, BigInteger.Compare(BigInteger.MinusOne, cor.NewBigInteger(-10))
    Assert.AreEqual -1, BigInteger.Compare(cor.NewBigInteger(-10), BigInteger.MinusOne)
End Sub

Public Sub Minus128AndMinus129()
    Dim Neg128 As BigInteger
    Dim Neg129 As BigInteger
    
    Set Neg128 = cor.NewBigInteger(-128)
    Set Neg129 = cor.NewBigInteger(-129)
    
    Assert.AreEqual 1, BigInteger.Compare(Neg128, Neg129)
End Sub

Public Sub CompareTwoEmpties()
    Dim b1 As New BigInteger
    Dim b2 As New BigInteger

    Assert.AreEqual 0, BigInteger.Compare(b1, b2)
    Assert.AreEqual 0, BigInteger.Compare(b2, b1)
    Assert.AreEqual 0, b1.CompareTo(b2)
    Assert.AreEqual 0, b2.CompareTo(b1)
End Sub

Public Sub CompareEquals()
    Dim b1 As BigInteger
    Set b1 = cor.NewBigInteger(12345)

    Dim b2 As BigInteger
    Set b2 = cor.NewBigInteger(12345)

    Assert.AreEqual 0, b1.CompareTo(b2)
    Assert.AreEqual 0, b2.CompareTo(b1)
End Sub

Public Sub CompareNotEquals()
    Dim b1 As BigInteger
    Set b1 = cor.NewBigInteger(12345)

    Dim b2 As BigInteger
    Set b2 = cor.NewBigInteger(54321)

    Assert.AreEqual -1, b1.CompareTo(b2)
    Assert.AreEqual 1, b2.CompareTo(b1)
    Assert.IsFalse b1.Equals(b2)
    Assert.IsFalse b2.Equals(b1)
End Sub

Public Sub CompareNegatives()
    Dim b1 As BigInteger
    Set b1 = cor.NewBigInteger(-1)

    Dim b2 As BigInteger
    Set b2 = cor.NewBigInteger(&H80000000)

    Assert.AreEqual -1, b2.CompareTo(b1)
    Assert.AreEqual 1, b1.CompareTo(b2)
End Sub

Public Sub CompareMixed()
    Dim b1 As BigInteger
    Set b1 = cor.NewBigInteger(-1)

    Dim b2 As BigInteger
    Set b2 = cor.NewBigInteger(2)

    Assert.AreEqual -1, b1.CompareTo(b2)
    Assert.AreEqual 1, b2.CompareTo(b1)
End Sub

Public Sub CompareTest1()
    Dim b1 As BigInteger
    Set b1 = cor.NewBigInteger(-65535)
    
    Dim b2 As BigInteger
    Set b2 = cor.NewBigInteger(-65536)
    
    Assert.AreEqual 1, b1.CompareTo(b2)
    Assert.AreEqual -1, b2.CompareTo(b1)
End Sub

Public Sub CompareTest2()
    Dim b1 As BigInteger
    Set b1 = cor.NewBigInteger(-65535)
    
    Dim b2 As BigInteger
    Set b2 = cor.NewBigInteger(-128)
    
    Assert.AreEqual -1, b1.CompareTo(b2)
    Assert.AreEqual 1, b2.CompareTo(b1)
End Sub


