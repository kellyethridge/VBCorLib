VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestNumberFormatInfoFlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestNumberFormatInfoFlt
'
Option Explicit

Private Info As NumberFormatInfo

Public Sub Setup()
    Set Info = New NumberFormatInfo
End Sub

Public Sub TestZero()
    Assert.AreEqual "0", Info.Format(0#)
End Sub

Public Sub TestGeneral()
    Assert.AreEqual "12345", Info.Format(12345#)
    Assert.AreEqual "-12345", Info.Format(-12345#)
    Assert.AreEqual "1.2345", Info.Format(1.2345)
    Assert.AreEqual "0.012345", Info.Format(0.012345)
    Assert.AreEqual "1234.5678", Info.Format(CVar(1234.5678))
    Assert.AreEqual "1234500000000000000000000000", Info.Format(1.2345E+27, "g28")
    Assert.AreEqual "1.2345E-05", Info.Format(0.000012345)
    Assert.AreEqual "1.2345E+15", Info.Format(1.2345E+15)
    Assert.AreEqual "1.235e+04", Info.Format(12345#, "g4")
    Assert.AreEqual "0", Info.Format(0#, "g")
End Sub

Public Sub TestFixed()
    Assert.AreEqual "12345.00", Info.Format(12345#, "f")
    Assert.AreEqual "12345", Info.Format(12345#, "f0")
    Assert.AreEqual "123450.0000000000", Info.Format(123450#, "f10")
    Assert.AreEqual "123.45", Info.Format(123.45, "f")
    Assert.AreEqual "1.23", Info.Format(1.2345, "f")
    Assert.AreEqual "1.235", Info.Format(1.2345, "f3")
    Assert.AreEqual "0.12", Info.Format(0.12345, "f")
    Assert.AreEqual "0.000", Info.Format(0.00012345, "f3")
    Assert.AreEqual "0.0001235", Info.Format(0.00012345, "f7")
    Assert.AreEqual "0.00012345000000000000", Info.Format(0.00012345, "f20")
    Assert.AreEqual "1.00", Info.Format(1#, "f")
    Assert.AreEqual "0.00", Info.Format(0#, "f")
End Sub

Public Sub TestScientific()
    Assert.AreEqual "1.234500e+004", Info.Format(12345#, "e")
    Assert.AreEqual "-1.2345e+004", Info.Format(-12345#, "e4")
    Assert.AreEqual "1.235e+004", Info.Format(12345#, "e3")
    Assert.AreEqual "1e+004", Info.Format(12345#, "e0")
    Assert.AreEqual "1.2E+004", Info.Format(12345#, "E1")
    Assert.AreEqual "1.234500e+006", Info.Format(1234500#, "e")
    Assert.AreEqual "1.235e+006", Info.Format(1234500#, "e3")
    Assert.AreEqual "1.234500e-003", Info.Format(0.0012345, "e")
    Assert.AreEqual "1.235e-003", Info.Format(0.0012345, "e3")
    Assert.AreEqual "1.2345000e-200", Info.Format(1.2345E-200, "e7")
    Assert.AreEqual "1.000000e+000", Info.Format(1#, "e")
    Assert.AreEqual "1.0e+000", Info.Format(1#, "e1")
    Assert.AreEqual "0.000000e+000", Info.Format(0#, "e")
    Assert.AreEqual "1.234e-001", Info.Format(0.1234, "e3")
End Sub

Public Sub TestDecimal()
    On Error GoTo errTrap
    Info.Format 1#, "d"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is FormatException, "Wrong exception type caught."
End Sub

Public Sub TestNumber()
    Assert.AreEqual "1.23", Info.Format(1.23, "n")
    Assert.AreEqual "1,234.56", Info.Format(1234.56, "n")
    Assert.AreEqual "1,234,567.89", Info.Format(1234567.89, "n")
    Assert.AreEqual "12,345,678.91", Info.Format(12345678.91, "n")
    Assert.AreEqual "123", Info.Format(123.45, "n0")
    Assert.AreEqual "123.45678000", Info.Format(123.45678, "n8")
    Assert.AreEqual "-123", Info.Format(-123.45, "n0")
    Assert.AreEqual "0.00", Info.Format(0.00123, "n")
    Assert.AreEqual "0.00124", Info.Format(0.001235, "n5")
    Assert.AreEqual "123.00", Info.Format(123#, "n")
    Assert.AreEqual "123,123,123,000,000,000,000.00000", Info.Format(1.23123123E+20, "n5")
    Assert.AreEqual "0.00", Info.Format(0#, "n")
    
    Info.NumberNegativePattern = [- n]
    Assert.AreEqual "- 123.00", Info.Format(-123#, "n")

    Info.NumberNegativePattern = [n-]
    Assert.AreEqual "123.00-", Info.Format(-123#, "n")

    Info.NumberNegativePattern = [n -]
    Assert.AreEqual "123.00 -", Info.Format(-123#, "n")

    Info.NumberNegativePattern = [(n)]
    Assert.AreEqual "(123.00)", Info.Format(-123#, "n")

    Dim sizes() As Long
    sizes = NewLongs(1, 2, 3)
    Info.NumberGroupSizes = sizes
    Assert.AreEqual "12,345,67,8.91", Info.Format(12345678.91, "n")

    sizes = NewLongs(1, 2, 1, 0)
    Info.NumberGroupSizes = sizes
    Assert.AreEqual "12345,6,78,9.12", Info.Format(123456789.12, "n")
    
    Info.NumberGroupSizes = cArray.CreateInstance(vbLong)
    Assert.AreEqual "1234567890", Info.Format(1234567890#, "n0")
End Sub

Public Sub TestCurrency()
    Assert.AreEqual "¤1.23", Info.Format(1.23, "c")
    Assert.AreEqual "¤1,234.56", Info.Format(1234.56, "c")
    Assert.AreEqual "¤1,234,567.89", Info.Format(1234567.89, "c")
    Assert.AreEqual "¤12,345,678.91", Info.Format(12345678.91, "c")
    Assert.AreEqual "¤123", Info.Format(123.45, "c0")
    Assert.AreEqual "¤123.45678000", Info.Format(123.45678, "c8")
    Assert.AreEqual "(¤123)", Info.Format(-123.45, "c0")
    Assert.AreEqual "¤0.00", Info.Format(0.00123, "c")
    Assert.AreEqual "¤0.00124", Info.Format(0.001235, "c5")
    Assert.AreEqual "¤123.00", Info.Format(123#, "c")
    Assert.AreEqual "¤123,123,123,000,000,000,000.00000", Info.Format(1.23123123E+20, "c5")
    Assert.AreEqual "¤0.00", Info.Format(0#, "c")
    
    Info.CurrencyNegativePattern = [-$ n]
    Assert.AreEqual "-¤ 123.00", Info.Format(-123#, "c")

    Info.CurrencyNegativePattern = [$ n-]
    Assert.AreEqual "¤ 123.00-", Info.Format(-123#, "c")

    Info.CurrencyNegativePattern = [(n $)]
    Assert.AreEqual "(123.00 ¤)", Info.Format(-123#, "c")

    Info.CurrencyNegativePattern = [($n)]
    Assert.AreEqual "(¤123.00)", Info.Format(-123#, "c")

    Info.CurrencyGroupSizes = NewLongs(1, 2, 3)
    Assert.AreEqual "¤12,345,67,8.91", Info.Format(12345678.91, "c")

    Info.CurrencyGroupSizes = NewLongs(1, 2, 1, 0)
    Assert.AreEqual "¤12345,6,78,9.12", Info.Format(123456789.12, "c")
    
    Info.CurrencyGroupSizes = NewLongs()
    Assert.AreEqual "¤1234567890", Info.Format(1234567890#, "c0")
End Sub

Public Sub TestPercent()
    Assert.AreEqual "0.00 %", Info.Format(0#, "p")
    Assert.AreEqual "123.00 %", Info.Format(1.23, "p")
    Assert.AreEqual "123,456.00 %", Info.Format(1234.56, "p")
    Assert.AreEqual "123,456,789.00 %", Info.Format(1234567.89, "p")
    Assert.AreEqual "1,234,567,891.00 %", Info.Format(12345678.91, "p")
    Assert.AreEqual "12,345 %", Info.Format(123.45, "p0")
    Assert.AreEqual "12,345.67800000 %", Info.Format(123.45678, "p8")
    Assert.AreEqual "-12,345 %", Info.Format(-123.45, "p0")
    Assert.AreEqual "0.12 %", Info.Format(0.00123, "p")
    Assert.AreEqual "0.124 %", Info.Format(0.001235, "p3")
    Assert.AreEqual "12,300.00 %", Info.Format(123#, "p")
    Assert.AreEqual "12,312,312,300,000,000,000,000.00000 %", Info.Format(1.23123123E+20, "p5")
    
    Info.PercentNegativePattern = [-# %]
    Assert.AreEqual "-12,300.00 %", Info.Format(-123#, "p")

    Info.PercentNegativePattern = [-#%]
    Assert.AreEqual "-12,300.00%", Info.Format(-123#, "p")

    Info.PercentNegativePattern = [-%#]
    Assert.AreEqual "-%12,300.00", Info.Format(-123#, "p")

    Info.PercentGroupSizes = NewLongs(1, 2, 3)
    Assert.AreEqual "1,234,567,89,1.00 %", Info.Format(12345678.91, "p")

    Info.PercentGroupSizes = NewLongs(1, 2, 1, 0)
    Assert.AreEqual "1234567,8,91,2.00 %", Info.Format(123456789.12, "p")
    
    Info.PercentGroupSizes = cArray.CreateInstance(vbLong)
    Assert.AreEqual "123456789000 %", Info.Format(1234567890#, "p0")
End Sub

Public Sub TestNAN()
    Dim d As Double
    On Error Resume Next
    d = 0# / 0#
    On Error GoTo 0
    Assert.AreEqual "NaN", Info.Format(d)
End Sub

Public Sub TestPosInfinity()
    Dim d As Double
    On Error Resume Next
    d = 1# / 0#
    On Error GoTo 0
    Assert.AreEqual "Infinity", Info.Format(d)
End Sub

Public Sub TestNegInfinity()
    Dim d As Double
    On Error Resume Next
    d = -1# / 0#
    On Error GoTo 0
    Assert.AreEqual "-Infinity", Info.Format(d)
End Sub
