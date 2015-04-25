VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBSubtractionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBSubtractionTests
'
Option Explicit
Implements ICategorizable

Private b As BigInteger

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Numerics"
End Sub

Public Sub Setup()
    Set b = Nothing
End Sub

Public Sub SubtractZeros()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(Nothing, Nothing)
    Assert.IsNotNothing b
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub SubtractOneFromZero()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(Nothing, BigInteger.One)
    Assert.IsNotNothing b
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-1", b.ToString
End Sub

Public Sub SubtractMinusOneFromZero()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(BigInteger.Zero, BigInteger.MinusOne)
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub SubtractTwoNegatives()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(BigInteger.MinusOne, BigInteger.MinusOne)
    Assert.AreEqual 0, b.Sign, "Sign should be 0"
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub Subtract10FromMinus120()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(Cor.NewBigInteger(-120), Cor.NewBigInteger(10))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-130", b.ToString
End Sub

Public Sub SubtractLargeNegatives()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(BigInteger.Parse("-1234567890"), BigInteger.Parse("-7788443322"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "6553875432", b.ToString
End Sub

Public Sub SubtractDifferentSigns()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(Cor.NewBigInteger(-123456), Cor.NewBigInteger(543))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual CStr(-123456 - 543), b.ToString
End Sub

Public Sub SubtractDifferentSigns2()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(Cor.NewBigInteger(1), Cor.NewBigInteger(-123456))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual CStr(1 - (-123456)), b.ToString
End Sub

Public Sub SubtractDifferentSigns3()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(Cor.NewBigInteger(-543), Cor.NewBigInteger(123456))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual CStr(-543 - (123456)), b.ToString
End Sub

Public Sub SubtractDifferentSigns4()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(Cor.NewBigInteger(123456), Cor.NewBigInteger(-3))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual CStr(123456 - (-3)), b.ToString
End Sub

Public Sub SubtractOne()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(BigInteger.Parse("1000000000000000"), BigInteger.One)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "999999999999999", b.ToString
End Sub

Public Sub SubtractMinusOne()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(BigInteger.Parse("999999999999999"), BigInteger.MinusOne)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1000000000000000", b.ToString
End Sub

Public Sub SubtractFromMinusOne()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(BigInteger.MinusOne, BigInteger.Parse("999999999999999"))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-1000000000000000", b.ToString
End Sub

Public Sub SubtractionLoop()
    Dim i As Long
    Dim b As BigInteger
    For i = 1 To 13
        Set b = BigInteger.Subtract(b, BigInteger.Ten)
    Next i

    Assert.AreEqual "-130", b.ToString
End Sub

Public Sub SubtractNegativeLoop()
    Dim i As Long
    Dim b As BigInteger
    For i = 1 To 1000
        Set b = BigInteger.Subtract(b, Cor.NewBigInteger(-10))
    Next i

    Assert.AreEqual "10000", b.ToString
End Sub

Public Sub SubtractLargeNegativeFromLittle()
    Dim b As BigInteger
    Set b = BigInteger.Subtract(BInt(123), BInt("12345678901234567890"))
    Assert.AreEqual "-12345678901234567767", b.ToString
End Sub


