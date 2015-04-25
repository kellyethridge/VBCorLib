VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBAdditionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' AdditionTests
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

Public Sub AddNothingsShouldEqual0()
    Set b = BigInteger.Add(Nothing, Nothing)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub AddOnes()
    Dim b As BigInteger
    Set b = BigInteger.Add(BigInteger.One, BigInteger.One)
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "2", b.ToString
End Sub

Public Sub AddThousands()
    Dim b As BigInteger
    Set b = BigInteger.Add(Cor.NewBigInteger(1234), Cor.NewBigInteger(4321))
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual CStr(1234 + 4321), b.ToString
End Sub

Public Sub AddDifferentLengths()
    Dim b As BigInteger
    Set b = BigInteger.Add(Cor.NewBigInteger(12), Cor.NewBigInteger(1234567))
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual CStr(12 + 1234567), b.ToString
End Sub

Public Sub AddNegatives()
    Dim b As BigInteger
    Set b = BigInteger.Add(BigInteger.MinusOne, BigInteger.MinusOne)
    Assert.IsNotNothing b
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-2", b.ToString
End Sub

Public Sub AddZeroAndOne()
    Dim b As BigInteger
    Set b = BigInteger.Add(BigInteger.Zero, BigInteger.One)
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub AddZeroAndMinusOne()
    Dim b As BigInteger
    Set b = BigInteger.Add(BigInteger.Zero, BigInteger.MinusOne)
    Assert.IsNotNothing b
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-1", b.ToString
End Sub

Public Sub AddOppositeSigns()
    Dim b As BigInteger
    Set b = BigInteger.Add(BigInteger.One, BigInteger.MinusOne)
    Assert.IsNotNothing b
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub AddLargeNegatives()
    Dim b As BigInteger
    Set b = BigInteger.Add(Cor.NewBigInteger(-123456), Cor.NewBigInteger(-5555222))
    Assert.IsNotNothing b
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual CStr(-123456 + (-5555222)), b.ToString
End Sub

Public Sub AddLargeOpposites()
    Dim b As BigInteger
    Set b = BigInteger.Add(Cor.NewBigInteger(123456), Cor.NewBigInteger(-5555222))
    Assert.IsNotNothing b
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual CStr(123456 + (-5555222)), b.ToString
End Sub

Public Sub AddLargeOpposites2()
    Dim b As BigInteger
    Set b = BigInteger.Add(Cor.NewBigInteger(-123456), Cor.NewBigInteger(5555222))
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual CStr(-123456 + (5555222)), b.ToString
End Sub

Public Sub AddLargeOpposites3()
    Dim b As BigInteger
    Set b = BigInteger.Add(Cor.NewBigInteger(5555222), Cor.NewBigInteger(-123))
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual CStr(-123 + (5555222)), b.ToString
End Sub

Public Sub ZeroPlusPositive()
    Dim b As BigInteger
    Set b = BigInteger.Add(Nothing, Cor.NewBigInteger(123456))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "123456", b.ToString
End Sub

Public Sub ZeroPlusNegative()
    Dim b As BigInteger
    Set b = BigInteger.Add(Nothing, Cor.NewBigInteger(-123456))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-123456", b.ToString
End Sub

Public Sub PositiveOverflow()
    Dim b As BigInteger
    Set b = BigInteger.Add(Cor.NewBigInteger(128), Cor.NewBigInteger(128))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "256", b.ToString
End Sub

Public Sub PositiveOverflow2()
    Dim b As BigInteger
    Set b = BigInteger.Add(Cor.NewBigInteger(64), Cor.NewBigInteger(64))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "128", b.ToString
End Sub

Public Sub AdditionLoop()
    Dim i As Long
    Set b = Cor.NewBigInteger(3)
    For i = 1 To 100
        Set b = b.Add(b)
    Next i
    
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "3802951800684688204490109616128", b.ToString
End Sub

Public Sub NegativeAdditionLoop()
    Dim i As Long
    Set b = Cor.NewBigInteger(-3)
    For i = 1 To 100
        Set b = b.Add(b)
    Next i
    
    Assert.AreEqual "-3802951800684688204490109616128", b.ToString
    Assert.AreEqual -1, b.Sign
End Sub

Public Sub AddLargeNegativeToLittle()
    Dim b As BigInteger
    Set b = BigInteger.Add(BInt(123), BInt("-12345678901234567890"))
    Assert.AreEqual "-12345678901234567767", b.ToString
End Sub



