VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBUnaryTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBUnaryTests
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

Public Sub NegateOfZero()
    Set b = BigInteger.Negate(BigInteger.Zero)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub NegateOfOne()
    Set b = BigInteger.Negate(BigInteger.One)
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-1", b.ToString
End Sub

Public Sub NegateOfMinusOne()
    Set b = BigInteger.Negate(BigInteger.MinusOne)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub NegateOfNothing()
    Set b = BigInteger.Negate(Nothing)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub Negate255()
    Set b = Cor.NewBigInteger(255)
    Set b = b.Negate
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual "-255", b.ToString
End Sub

Public Sub Negate00FFFF()
    Set b = NewBigInteger(65535)
    Set b = b.Negate
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual NewBytes(1, 0, &HFF, &HFF), b.ToArray
End Sub

Public Sub Negate65534()
    Set b = NewBigInteger(65534)
    Set b = b.Negate
    Assert.AreEqual -1, b.Sign
    Assert.AreEqual NewBytes(2, 0, &HFF, &HFF), b.ToArray
End Sub

Public Sub AbsoluteOfZero()
    Set b = BigInteger.Absolute(BigInteger.Zero)
    Assert.AreEqual 0, b.Sign
End Sub

Public Sub AbsoluteOfMinusOne()
    Set b = BigInteger.Absolute(BigInteger.MinusOne)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

Public Sub AbsoluteOfNothing()
    Set b = BigInteger.Absolute(Nothing)
    Assert.AreEqual 0, b.Sign
    Assert.AreEqual "0", b.ToString
End Sub

Public Sub AbsolueOfOne()
    Set b = BigInteger.Absolute(BigInteger.One)
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "1", b.ToString
End Sub

