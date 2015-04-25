VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBPowTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBPowTests
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

Public Sub ZeroToZero()
    Set b = BigInteger.Pow(BigInteger.Zero, BigInteger.Zero)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub ZeroToOne()
    Set b = BigInteger.Pow(BigInteger.Zero, BigInteger.One)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub TwoToOne()
    Set b = BigInteger.Pow(BInt(2), BigInteger.One)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "2", b.ToString
End Sub

Public Sub TwoToTwo()
    Set b = BigInteger.Pow(BInt(2), BInt(2))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "4", b.ToString
End Sub

Public Sub TwoToFour()
    Set b = BigInteger.Pow(BInt(2), BInt(4))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "16", b.ToString
End Sub

Public Sub MinusTwoToTwo()
    Set b = BigInteger.Pow(BInt(-2), BInt(2))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "4", b.ToString
End Sub

Public Sub MinusTwoToThree()
    Set b = BigInteger.Pow(BInt(-2), BInt(3))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-8", b.ToString
End Sub

Public Sub BigPow()
    Set b = BigInteger.Pow(BInt(-987654321), BInt(11))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-872277458049782690634939797742584107247797490258292883792361746677987442093642323779216637616549521", b.ToString
End Sub

Public Sub ReallyBigPow()
    Set b = BigInteger.Pow(BInt(-987), BInt(117))
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-216325074696309558783776376303518372118577753788437518592642281696571828814743344779689654716301205610596884621347900456103979399101884771627786428610000978709856040709278643367625707116724052133903489423814892947941267761644861023184115121424748190327368213823471381999178537166830777594099715872466290716294812956953795520066705650738011057200857067", b.ToString
End Sub

Public Sub Pow2()
    Set b = BigInteger.Pow(BInt(256), BInt(256))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1" & String$(512, 48), b.ToString("X")
End Sub

Public Sub ModPow1()
    Set b = BigInteger.ModPow(BInt(2), BInt(4), BInt(6))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "4", b.ToString
End Sub

Public Sub ModPow2()
    Set b = BigInteger.ModPow(BInt(256), BInt(256), BInt(45911))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "8099", b.ToString
End Sub
