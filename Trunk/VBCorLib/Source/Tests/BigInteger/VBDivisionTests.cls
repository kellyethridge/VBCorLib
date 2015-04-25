VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBDivisionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBDivisionTests
'
Option Explicit
Implements ICategorizable

Private b As BigInteger
Private r As BigInteger

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub Setup()
    Set b = Nothing
    Set r = Nothing
End Sub

Public Sub DivideByZeroError()
    On Error GoTo errTrap
    BigInteger.Divide Nothing, Nothing
    Assert.Fail "An error should occur."
    
errTrap:
    Assert.AreEqual 11, Err.Number
End Sub

Public Sub DivideZeroReturnsZero()
    Set b = BigInteger.Divide(BigInteger.Zero, BigInteger.One)
    Assert.IsNotNothing b
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub DivideByOneReturnsDividend()
    Set b = BigInteger.Divide(BigInteger.Ten, BigInteger.One)
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "10", b.ToString
End Sub

Public Sub DivideByMinusOneReturnsNegative()
    Set b = BigInteger.Divide(BigInteger.Ten, BigInteger.MinusOne)
    Assert.IsNotNothing b
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-10", b.ToString
End Sub

Public Sub DivideNegativeByNegativeReturnsPositive()
    Set b = BigInteger.Divide(Cor.NewBigInteger(-10), BigInteger.MinusOne)
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "10", b.ToString
End Sub

Public Sub SpecialNumber()
    Dim n1 As BigInteger
    Dim n2 As BigInteger
    
    Set n1 = BigInteger.Parse("170141183420855150493001878992821682176")
    Set n2 = BigInteger.Parse("39614081266355540842216685573")
    
    Set b = BigInteger.DivRem(n1, n2, r)
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "4294967293", b.ToString
    Assert.IsNotNothing r
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "39614081266355540837921718287", r.ToString
End Sub

Public Sub SpecialNegativeNumber()
    Dim n1 As BigInteger
    Dim n2 As BigInteger

    Set n1 = BigInteger.Parse("-170141183420855150493001878992821682176")
    Set n2 = BigInteger.Parse("39614081266355540842216685573")
    
    Set b = BigInteger.DivRem(n1, n2, r)
    Assert.IsNotNothing b
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-4294967293", b.ToString
    Assert.IsNotNothing r
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-39614081266355540837921718287", r.ToString
End Sub

Public Sub Negative10000By300()
    Set b = BigInteger.DivRem(BInt(-10000), BInt(300), r)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-33", b.ToString
    Assert.AreEqual -1, r.Sign
    Assert.AreEqual "-100", r.ToString
End Sub

Public Sub Divide2()
    Set b = BigInteger.DivRem(BInt(646), BInt(68), r)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "9", b.ToString
    Assert.AreEqual 1, r.Sign
    Assert.AreEqual "34", r.ToString
End Sub

Public Sub Divide3()
    Set b = BigInteger.DivRem(BInt(8134), BInt(7814), r)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
    Assert.AreEqual 1, r.Sign
    Assert.AreEqual "320", r.ToString
End Sub

Public Sub Divide4()
    Set b = BigInteger.DivRem(BInt(1234), BInt(3), r)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "411", b.ToString
    Assert.AreEqual 1, r.Sign
    Assert.AreEqual "1", r.ToString
End Sub

Public Sub Divide5()
    Set b = BigInteger.DivRem(BInt(40902), BInt(24140), r)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
    Assert.AreEqual 1, r.Sign
    Assert.AreEqual "16762", r.ToString
End Sub

Public Sub Divide6()
    Dim r As BigInteger
    Set b = BigInteger.DivRem(BInt("-9876543210987654321"), BInt(11), r)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-897867564635241301", b.ToString
    Assert.AreEqual -1, r.Sign
    Assert.AreEqual "-10", r.ToString
End Sub

Public Sub Divide7()
    Dim r As BigInteger
    Set b = BigInteger.DivRem(BInt("0x00000000ffffffffffffffffffffffffffffffff"), BInt("0x00000000ffffffff"), r)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "79228162532711081671548469249", b.ToString
    Assert.AreEqual 0, r.Sign
End Sub

Public Sub Divide8()
    Dim r As BigInteger
    Set b = BigInteger.DivRem(BInt("460703086179019821693043897409119083541759798939670205282695248745968960462491099024560450123575"), BInt("210830276673471"), r)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "2185184658712685669515287106463728011552413047504880360572766918118852395516806406", b.ToString
    Assert.AreEqual "147913267068349", r.ToString
End Sub

Public Sub Divide9()
    Set b = BigInteger.Divide(BInt("0x0000ffffffffffffffffffff"), BInt("0x0000ffffffff"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1000000010000", b.ToString("x")
End Sub

Public Sub Divide10()
    Set b = BigInteger.Divide(BInt("0x0000EEEEEEEEEEEEEEEEEEEE"), BInt("0x0000ffff"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "EEEFDDDECCCDBBBC", b.ToString("X")
End Sub

Public Sub Divide11()
    Set b = BigInteger.Divide(BInt("0x0000EEEEEEEEEEEEEEEEEEEE"), BInt("0x0000EEEE"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "10001000100010001", b.ToString("X")
End Sub

Public Sub Divide12()
    Set b = BigInteger.Divide(BInt("0x0000EEEEEEEEEEEEEEEEEEEE"), BInt("0x0000EEEEEEEE"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1000000010000", b.ToString("X")
End Sub

Public Sub Divide13()
    Dim r As BigInteger
    Set b = BigInteger.DivRem(BInt("0x0000EEEEEEEEEEEEEEEEEEEE"), BInt("0x0001eeee"), r)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "7B964D51631E7678", b.ToString("X")
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "13B5E", r.ToString("X")
End Sub

Public Sub Divide14()
    Dim r As BigInteger
    Set b = BigInteger.DivRem(BInt("0x0000EEEEEEEEEEEEEEEEEEEE"), BInt("0x00010001"), r)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "EEEE0000EEEE0000", b.ToString("X")
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "EEEE", r.ToString("X")
End Sub

Public Sub Divide15()
    Dim r As BigInteger
    Set b = BigInteger.DivRem(BInt("0x0000ffffffffffffffffffff"), BInt("0x00010001"), r)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "ffff0000ffff0000", b.ToString("x")
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "ffff", r.ToString("x")
End Sub

Public Sub Divide16()
    Dim r As BigInteger
    Set b = BigInteger.DivRem(BInt("0x10000000000000000"), BInt(45911), r)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "401793558705093", b.ToString
    Assert.AreEqual "26893", r.ToString
End Sub

Public Sub GCD1()
    Set b = BigInteger.GreatestCommonDivisor(BInt(646), BInt(68))
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "34", b.ToString
End Sub

Public Sub GCD2()
    Set b = BigInteger.GreatestCommonDivisor(BInt(40902), BInt(24140))
    Assert.IsNotNothing b
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "34", b.ToString
End Sub

Public Sub DivideSmallByLarge()
    Set b = BigInteger.DivRem(BInt("9999999"), BInt("10000000"), r)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
    Assert.AreEqual 1, r.Sign
    Assert.AreEqual "9999999", r.ToString
End Sub

Public Sub DivideSmallByLarge2()
    Set b = BigInteger.DivRem(BInt("&h0000ffff"), BInt("&h00ffffff"), r)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
    Assert.AreEqual 1, r.Sign
    Assert.AreEqual "ffff", r.ToString("x")
End Sub

Public Sub DivideSmallByLarge3()
    Set b = BigInteger.DivRem(BInt(-12), BInt(999999), r)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
    Assert.AreEqual -1, r.Sign
    Assert.AreEqual "-12", r.ToString
End Sub


