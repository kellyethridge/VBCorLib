VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBParseDecimalTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBParseDecimalTests
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


Public Sub TryEmptyString()
    Assert.IsFalse BigInteger.TryParse("", b)
    Assert.IsNothing b
End Sub

Public Sub Try0()
    Assert.IsTrue BigInteger.TryParse("0", b)
    AssertNumber 0, 0, Zero, NewBytes
End Sub

Public Sub Try1()
    Assert.IsTrue BigInteger.TryParse("1", b)
    AssertNumber 1, 2, Positive, NewBytes(1)
End Sub

Public Sub Try01()
    Assert.IsTrue BigInteger.TryParse("01", b)
    AssertNumber 1, 2, Positive, NewBytes(1)
End Sub

Public Sub TryNegative1()
    Assert.IsTrue BigInteger.TryParse("-1", b)
    AssertNumber 2, 2, Negative, NewBytes(&HFF, &HFF)
End Sub

Public Sub Try65535()
    Assert.IsTrue BigInteger.TryParse("65535", b)
    AssertNumber 2, 3, Positive, NewBytes(&HFF, &HFF)
End Sub

Public Sub TryInvalidCharacter()
    Assert.IsFalse BigInteger.TryParse("1234.134", b)
    Assert.IsNothing b
End Sub

Public Sub TryWhiteSpace()
    Assert.IsTrue BigInteger.TryParse("   65536    ", b)
    AssertNumber 3, 4, Positive, NewBytes(0, 0, &H1)
End Sub

Public Sub TryWhiteSpaceNegative()
    Assert.IsTrue BigInteger.TryParse("    -65536   ", b)
    AssertNumber 4, 4, Negative, NewBytes(0, 0, &HFF, &HFF)
End Sub

Public Sub TryLeadingZeros()
    Assert.IsTrue BigInteger.TryParse("   000000000000000000000255    ", b)
    AssertNumber 1, 2, Positive, NewBytes(255)
End Sub

Public Sub TryLargePositive()
    Assert.IsTrue BigInteger.TryParse("12345678901234567890", b)
    AssertNumber 8, 9, Positive, NewBytes(&HD2, &HA, &H1F, &HEB, &H8C, &HA9, &H54, &HAB)
End Sub

Public Sub TryLargeNegative()
    Assert.IsTrue BigInteger.TryParse("-12345678901234567890", b)
    AssertNumber 10, 9, Negative, NewBytes(&H2E, &HF5, &HE0, &H14, &H73, &H56, &HAB, &H54, &HFF, &HFF)
End Sub

Public Sub ParseLargeNegative()
    Set b = BigInteger.Parse("-12345678901234567890")
    AssertNumber 10, 9, Negative, NewBytes(&H2E, &HF5, &HE0, &H14, &H73, &H56, &HAB, &H54, &HFF, &HFF)
End Sub

Public Sub FailToParse()
    Assert.Throws 5
    Set b = BigInteger.Parse("-12345678901234567890a")
End Sub



Private Sub AssertNumber(ByVal Precision As Long, ByVal Length As Long, ByVal Sign As Sign, ByRef Bytes() As Byte)
    Assert.AreEqual Precision, b.Precision, "Wrong Precision"
'    Assert.AreEqual Length, b.Length, "Wrong Length"
    Assert.AreEqual Sign, b.Sign, "Wrong Sign"
    Assert.AreEqual Bytes, b.ToArray, "Wrong Array"
End Sub



