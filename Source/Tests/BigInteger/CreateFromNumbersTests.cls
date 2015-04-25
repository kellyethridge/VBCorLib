VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBCreateFromNumbersTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CreateFromNumbersTests
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


Public Sub EmptyNumber()
    Set b = New BigInteger
    AssertNumber Zero, 0, NewBytes()
End Sub

Public Sub TestZero()
    Set b = NewBigInteger(0)
    AssertNumber Zero, 0, NewBytes()
End Sub

Public Sub TestPositiveOne()
    Set b = NewBigInteger(1)
    AssertNumber Positive, 1, NewBytes(1)
End Sub

Public Sub TestNegativeOne()
    Set b = NewBigInteger(&HFFFFFFFF)
    AssertNumber Negative, 2, NewBytes(&HFF, &HFF)
End Sub

Public Sub TestPositive8000()
    Set b = NewBigInteger(&H8000&)
    AssertNumber Positive, 2, NewBytes(0, &H80)
End Sub

Public Sub TestNegative8000()
    Set b = NewBigInteger(&H8000)
    AssertNumber Negative, 2, NewBytes(0, &H80)
End Sub

Public Sub TestFF8000()
    Set b = NewBigInteger(&HFF8000)
    AssertNumber Positive, 3, NewBytes(0, &H80, &HFF)
End Sub

Public Sub TestFFFF8000()
    Set b = NewBigInteger(&HFFFF8000)
    AssertNumber Negative, 2, NewBytes(0, &H80)
End Sub

Public Sub Test7FFF()
    Set b = NewBigInteger(&H7FFF)
    AssertNumber Positive, 2, NewBytes(&HFF, &H7F)
End Sub

Public Sub TestFF7FFF()
    Set b = NewBigInteger(&HFF7FFF)
    AssertNumber Positive, 3, NewBytes(&HFF, &H7F, &HFF)
End Sub

Public Sub TestFFFF7FFF()
    Set b = NewBigInteger(&HFFFF7FFF)
    AssertNumber Negative, 4, NewBytes(&HFF, &H7F, &HFF, &HFF)
End Sub

Public Sub TestFF()
    Set b = NewBigInteger(&HFF)
    AssertNumber Positive, 1, NewBytes(&HFF)
End Sub

Public Sub TestFloat1()
    Set b = NewBigInteger(1#)
    AssertNumber Positive, 1, NewBytes(1)
End Sub

Public Sub TestFloatNegative1()
    Set b = NewBigInteger(-1#)
    AssertNumber Negative, 2, NewBytes(&HFF, &HFF)
End Sub

Public Sub TestFloat1245Point6789()
    Set b = NewBigInteger(1245.6789)
    AssertNumber Positive, 2, NewBytes(&HDD, &H4)
End Sub

Public Sub TestFloat0()
    Set b = NewBigInteger(0#)
    AssertNumber Zero, 0, NewBytes()
End Sub

Public Sub TestFloatLessThan1()
    Set b = NewBigInteger(0.95)
    AssertNumber Zero, 0, NewBytes()
End Sub

Public Sub TestCurrency1()
    Set b = NewBigInteger(1@)
    AssertNumber Positive, 1, NewBytes(1)
End Sub

Public Sub TestCurrencyNegative1()
    Set b = NewBigInteger(-1@)
    AssertNumber Negative, 2, NewBytes(&HFF, &HFF)
End Sub

Public Sub TestDecimal1()
    Set b = NewBigInteger(CDec("1"))
    AssertNumber Positive, 1, NewBytes(1)
End Sub

Public Sub TestDecimalNegative1()
    Set b = NewBigInteger(CDec("-1"))
    AssertNumber Negative, 2, NewBytes(&HFF, &HFF)
End Sub

Public Sub TestDecimalWithFraction()
    Set b = NewBigInteger(CDec("12345678901231234567890123.456789"))
    Assert.AreEqual 1, b.Sign
    Assert.AreEqual "A364C98227BA2515FE4CB", b.ToString("X")
End Sub








Private Sub AssertNumber(ByVal Sign As Sign, ByVal Precision As Long, ByRef Arr() As Byte)
    Assert.IsNotNothing b, "b should not be Nothing"
    Assert.AreEqual Sign, b.Sign, "Wrong sign"
    Assert.AreEqual Precision, b.Precision, "Wrong precision"
    
    Dim a() As Byte
    a = b.ToArray
    Assert.AreEqual Arr, a, "Wrong array"
End Sub
