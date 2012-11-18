VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestNumberFormatInfoInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private info As NumberFormatInfo

Public Sub Setup()
    Set info = New NumberFormatInfo
End Sub


Public Sub TestZero()
    Assert.AreEqual "0", info.Format(0)
End Sub

Public Sub TestGeneral()
    Assert.AreEqual "12300", info.Format(12300, "")
    Assert.AreEqual "12345", info.Format(12345)
    Assert.AreEqual "1.23e+04", info.Format(12345, "g3")
    Assert.AreEqual "1.235e+04", info.Format(12345, "g4")
    Assert.AreEqual "1.23E+04", info.Format(12345, "G3")
    Assert.AreEqual "-12345", info.Format(-12345)
    Assert.AreEqual "2e+01", info.Format(15, "g1")
    Assert.AreEqual "12345", info.Format(12345, "g26")
    Assert.AreEqual "ggg", info.Format(12345, "ggg")
    Assert.AreEqual "12345", info.Format(12345, "g0")
    Assert.AreEqual "0", info.Format(0, "g")
End Sub

Public Sub TestFixed()
    Assert.AreEqual "12345.00", info.Format(12345, "f")
    Assert.AreEqual "-12345", info.Format(-12345, "f0")
    Assert.AreEqual "12345.0000000000", info.Format(12345, "f10")
    Assert.AreEqual "1234500.000", info.Format(1234500, "F3")
    Assert.AreEqual "fff", info.Format(12345, "fff")
    Assert.AreEqual "1.00", info.Format(1, "f")
    Assert.AreEqual "0.00", info.Format(0, "f")
End Sub

Public Sub TestScientific()
    Assert.AreEqual "1.234500e+004", info.Format(12345, "e")
    Assert.AreEqual "-1.2345e+004", info.Format(-12345, "e4")
    Assert.AreEqual "1.235e+004", info.Format(12345, "e3")
    Assert.AreEqual "1e+004", info.Format(12345, "e0")
    Assert.AreEqual "1.2E+004", info.Format(12345, "E1")
    Assert.AreEqual "1.000000e+000", info.Format(1, "e")
    Assert.AreEqual "1.0e+000", info.Format(1, "e1")
    Assert.AreEqual "1e+000", info.Format(1, "e0")
    Assert.AreEqual "0.000000e+000", info.Format(0, "e")
End Sub

Public Sub TestDecimal()
    Assert.AreEqual "12345", info.Format(12345, "d")
    Assert.AreEqual "12345", info.Format(12345, "d1")
    Assert.AreEqual "0012345", info.Format(12345, "d7")
    Assert.AreEqual "0", info.Format(0, "d")
End Sub

Public Sub TestHexidecimal()
    Assert.AreEqual "A0", info.Format(160, "X")
    Assert.AreEqual "ab", info.Format(171, "x")
    Assert.AreEqual "ffffffff", info.Format(CLng(-1), "x")
    Assert.AreEqual "ffff", info.Format(CInt(-1), "x")
    Assert.AreEqual "00123", info.Format(&H123, "x5")
    Assert.AreEqual "ABCDE", info.Format(&HABCDE, "X1")
    Assert.AreEqual "000", info.Format(0, "x3")
    Assert.AreEqual "ff", info.Format(CByte(255), "x")
End Sub

Public Sub TestNumber()
    Assert.AreEqual "0.00", info.Format(0, "n")
    Assert.AreEqual "1.00", info.Format(1, "n")
    Assert.AreEqual "1,234.00", info.Format(1234, "n")
    Assert.AreEqual "1,234,567.00", info.Format(1234567, "n")
    Assert.AreEqual "12,345,678.00", info.Format(12345678, "n")
    Assert.AreEqual "123", info.Format(123, "n0")
    Assert.AreEqual "123.00000", info.Format(123, "n5")
    Assert.AreEqual "-123", info.Format(-123, "n0")

    info.NumberNegativePattern = [- n]
    Assert.AreEqual "- 123.00", info.Format(-123, "n")

    info.NumberNegativePattern = [n-]
    Assert.AreEqual "123.00-", info.Format(-123, "n")

    info.NumberNegativePattern = [n -]
    Assert.AreEqual "123.00 -", info.Format(-123, "n")

    info.NumberNegativePattern = [(n)]
    Assert.AreEqual "(123.00)", info.Format(-123, "n")

    Dim sizes() As Long
    sizes = NewLongs(1, 2, 3)
    info.NumberGroupSizes = sizes
    Assert.AreEqual "12,345,67,8.00", info.Format(12345678, "n")

    sizes = NewLongs(1, 2, 1, 0)
    info.NumberGroupSizes = sizes
    Assert.AreEqual "12345,6,78,9.00", info.Format(123456789, "n")
    
    sizes = NewLongs(2, 3)
    info.NumberGroupSizes = sizes
    Assert.AreEqual "1.00", info.Format(1, "n")
    Assert.AreEqual "100,000,00.00", info.Format(10000000, "n")
    Assert.AreEqual "10,000,000,00.00", info.Format(1000000000, "n")
End Sub

Public Sub TestCurrency()
    Assert.AreEqual "¤0.00", info.Format(0, "c")
    Assert.AreEqual "¤1.00", info.Format(1, "c")
    Assert.AreEqual "¤1,234.00", info.Format(1234, "c")
    Assert.AreEqual "¤1,234,567.00", info.Format(1234567, "c")
    Assert.AreEqual "¤12,345,678.00", info.Format(12345678, "c")
    Assert.AreEqual "¤123", info.Format(123, "c0")
    Assert.AreEqual "¤123.00000", info.Format(123, "c5")
    Assert.AreEqual "(¤123)", info.Format(-123, "c0")
    
    info.CurrencyNegativePattern = [-$ n]
    Assert.AreEqual "-¤ 123.00", info.Format(-123, "c")
    
    info.CurrencyNegativePattern = [$ n-]
    Assert.AreEqual "¤ 123.00-", info.Format(-123, "c")
    
    info.CurrencyGroupSizes = NewLongs(1, 2, 3)
    Assert.AreEqual "¤12,345,67,8.00", info.Format(12345678, "c")
    
    info.CurrencyGroupSizes = NewLongs(1, 2, 1, 0)
    Assert.AreEqual "¤12345,6,78,9.00", info.Format(123456789, "c")
End Sub

Public Sub TestPercent()
    Assert.AreEqual "0.00 %", info.Format(0, "p")
    Assert.AreEqual "100.00 %", info.Format(1, "p")
    Assert.AreEqual "123,400.00 %", info.Format(1234, "p")
    Assert.AreEqual "123,456,700.00 %", info.Format(1234567, "p")
    Assert.AreEqual "1,234,567,800.00 %", info.Format(12345678, "p")
    Assert.AreEqual "12,300 %", info.Format(123, "p0")
    Assert.AreEqual "12,300.00000 %", info.Format(123, "p5")
    Assert.AreEqual "-12,300 %", info.Format(-123, "p0")
    
    info.PercentNegativePattern = [-# %]
    Assert.AreEqual "-12,300.00 %", info.Format(-123, "p")
    
    info.PercentNegativePattern = [-#%]
    Assert.AreEqual "-12,300.00%", info.Format(-123, "p")
    
    info.PercentNegativePattern = [-%#]
    Assert.AreEqual "-%12,300.00", info.Format(-123, "p")
    
    info.PercentGroupSizes = NewLongs(1, 2, 3)
    Assert.AreEqual "1,234,567,80,0.00 %", info.Format(12345678, "p")
    
    info.PercentGroupSizes = NewLongs(1, 2, 1, 0)
    Assert.AreEqual "1234567,8,90,0.00 %", info.Format(123456789, "p")
End Sub

