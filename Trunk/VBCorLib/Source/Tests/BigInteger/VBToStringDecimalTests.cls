VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBToStringDecimalTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBToStringDecimalTests
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


Public Sub TestZero()
    Set b = New BigInteger
    Assert.AreEqual "0", b.ToString
    Assert.AreEqual "0000", b.ToString("d4")
End Sub

Public Sub Test1To1000000000()
    Dim i As Long
    i = 1
    Do
        Set b = NewBigInteger(i)
        Assert.AreEqual CStr(i), b.ToString
        i = i * 10
    Loop While i < 1000000000
End Sub

Public Sub TestNegative1To1000000000()
    Dim i As Long
    i = 1
    Do
        Set b = NewBigInteger(-i)
        Assert.AreEqual CStr(-i), b.ToString
        i = i * 10
    Loop While i < 1000000000
End Sub

Public Sub TestCountLessThanSize()
    Set b = Cor.NewBigInteger(&H8000&)
    Assert.AreEqual "32768", b.ToString("d1")
End Sub

Public Sub TestCountGreaterThanSize()
    Set b = Cor.NewBigInteger(&H8000&)
    Assert.AreEqual "0000032768", b.ToString("d10")
End Sub

Public Sub TestCountEqualsSize()
    Set b = Cor.NewBigInteger(&H8000&)
    Assert.AreEqual "32768", b.ToString("d5")
End Sub

Public Sub TestCount1LargerThanSize()
    Set b = Cor.NewBigInteger(&H8000&)
    Assert.AreEqual "032768", b.ToString("d6")
End Sub

Public Sub TestNegativeCountLessThanSize()
    Set b = Cor.NewBigInteger(&H8000)
    Assert.AreEqual "-32768", b.ToString("d0")
End Sub

Public Sub TestNegativeCount1GreaterThanSize()
    Set b = Cor.NewBigInteger(&H8000)
    Assert.AreEqual "-032768", b.ToString("d6")
End Sub

Public Sub TestNegativeCountEqualsSize()
    Set b = Cor.NewBigInteger(&H8000)
    Assert.AreEqual "-32768", b.ToString("d5")
End Sub


