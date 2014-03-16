VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestNumberFormatInfoSng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestNumberFormatInfoSng
'
Option Explicit

Private Info As NumberFormatInfo

Public Sub Setup()
    Set Info = New NumberFormatInfo
End Sub

Public Sub TestZero()
    Assert.AreEqual "0", Info.Format(0!)
End Sub

'Public Sub TestGeneral()
'    Assert.AreEqual "12345", Info.Format(12345!)
'    Assert.AreEqual "-12345", Info.Format(-12345!)
'    Assert.AreEqual "1.2345", Info.Format(1.2345!)
'    Assert.AreEqual "0.012345", Info.Format(0.012345!)
'    Assert.AreEqual "1234.568", Info.Format(1234.568!)
'    Assert.AreEqual "12345000400000000000000000", Info.Format(1.2345E+25!, "g26")
'    Assert.AreEqual "1.2345E-05", Info.Format(1.2345E-05!)
'    Assert.AreEqual "1.2345E+15", Info.Format(1.2345E+15!)
'    Assert.AreEqual "1.235e+04", Info.Format(12345!, "g4")
'End Sub

Public Sub TestFixed()
    Assert.AreEqual "12345.00", Info.Format(12345!, "f")
    Assert.AreEqual "12345", Info.Format(12345!, "f0")
    Assert.AreEqual "123450.0000000000", Info.Format(123450!, "f10")
    Assert.AreEqual "123.45", Info.Format(123.45!, "f")
    Assert.AreEqual "1.23", Info.Format(1.2345!, "f")
    Assert.AreEqual "1.235", Info.Format(1.2345!, "f3")
    Assert.AreEqual "0.12", Info.Format(0.12345!, "f")
    Assert.AreEqual "0.000", Info.Format(1.2345E-04!, "f3")
    Assert.AreEqual "0.0001235", Info.Format(1.2345E-04!, "f7")
    Assert.AreEqual "0.00012345000000000000", Info.Format(1.2345E-04!, "f20")
End Sub

Public Sub TestScientific()
    Assert.AreEqual "1.234500e+004", Info.Format(12345!, "e")
    Assert.AreEqual "-1.2345e+004", Info.Format(-12345!, "e4")
    Assert.AreEqual "1.235e+004", Info.Format(12345!, "e3")
    Assert.AreEqual "1e+004", Info.Format(12345!, "e0")
    Assert.AreEqual "1.2E+004", Info.Format(12345!, "E1")
    Assert.AreEqual "1.234500e+006", Info.Format(1234500!, "e")
    Assert.AreEqual "1.235e+006", Info.Format(1234500!, "e3")
    Assert.AreEqual "1.234500e-003", Info.Format(0.0012345!, "e")
    Assert.AreEqual "1.235e-003", Info.Format(0.0012345!, "e3")
    Assert.AreEqual "1.2345000e-020", Info.Format(1.2345E-20!, "e7")
End Sub

Public Sub TestNumber()
    Assert.AreEqual "1.23", Info.Format(1.23!, "n")
    Assert.AreEqual "1,234.56", Info.Format(1234.56!, "n")
    Assert.AreEqual "1,234,568.00", Info.Format(1234568!, "n")
    Assert.AreEqual "12,345,680.00", Info.Format(1.234568E+07!, "n")
    Assert.AreEqual "123", Info.Format(123.45!, "n0")
    Assert.AreEqual "123.45680000", Info.Format(123.4568!, "n8")
    Assert.AreEqual "-123", Info.Format(-123.45!, "n0")
    Assert.AreEqual "0.00", Info.Format(0.00123!, "n")
    Assert.AreEqual "0.00124", Info.Format(0.001235!, "n5")
    Assert.AreEqual "123.00", Info.Format(123!, "n")
    Assert.AreEqual "123,123,100,000,000,000,000.00000", Info.Format(1.231231E+20!, "n5")
    
    Info.NumberNegativePattern = [- n]
    Assert.AreEqual "- 123.00", Info.Format(-123!, "n")

    Info.NumberNegativePattern = [n-]
    Assert.AreEqual "123.00-", Info.Format(-123!, "n")

    Info.NumberNegativePattern = [n -]
    Assert.AreEqual "123.00 -", Info.Format(-123!, "n")

    Info.NumberNegativePattern = [(n)]
    Assert.AreEqual "(123.00)", Info.Format(-123!, "n")

    Dim sizes() As Long
    sizes = NewLongs(1, 2, 3)
    Info.NumberGroupSizes = sizes
    Assert.AreEqual "12,345,68,0.00", Info.Format(1.234568E+07!, "n")

    sizes = NewLongs(1, 2, 1, 0)
    Info.NumberGroupSizes = sizes
    Assert.AreEqual "12345,6,80,0.00", Info.Format(1.234568E+08!, "n")
    
    Info.NumberGroupSizes = cArray.CreateInstance(vbLong)
    Assert.AreEqual "1234568000", Info.Format(1.234568E+09!, "n0")
End Sub

Public Sub TestCurrency()
    Assert.AreEqual "¤1.23", Info.Format(1.23!, "c")
    Assert.AreEqual "¤1,234.56", Info.Format(1234.56!, "c")
    Assert.AreEqual "¤1,234,568.00", Info.Format(1234568!, "c")
    Assert.AreEqual "¤12,345,680.00", Info.Format(1.234568E+07!, "c")
    Assert.AreEqual "¤123", Info.Format(123.45!, "c0")
    Assert.AreEqual "¤123.45680000", Info.Format(123.4568!, "c8")
    Assert.AreEqual "(¤123)", Info.Format(-123.45!, "c0")
    Assert.AreEqual "¤0.00", Info.Format(0.00123!, "c")
    Assert.AreEqual "¤0.00124", Info.Format(0.001235!, "c5")
    Assert.AreEqual "¤123.00", Info.Format(123!, "c")
    Assert.AreEqual "¤123,123,100,000,000,000,000.00000", Info.Format(1.231231E+20!, "c5")
    
    Info.CurrencyNegativePattern = [-$ n]
    Assert.AreEqual "-¤ 123.00", Info.Format(-123!, "c")

    Info.CurrencyNegativePattern = [$ n-]
    Assert.AreEqual "¤ 123.00-", Info.Format(-123!, "c")

    Info.CurrencyNegativePattern = [(n $)]
    Assert.AreEqual "(123.00 ¤)", Info.Format(-123!, "c")

    Info.CurrencyNegativePattern = [($n)]
    Assert.AreEqual "(¤123.00)", Info.Format(-123!, "c")

    Info.CurrencyGroupSizes = NewLongs(1, 2, 3)
    Assert.AreEqual "¤12,345,68,0.00", Info.Format(1.234568E+07!, "c")

    Info.CurrencyGroupSizes = NewLongs(1, 2, 1, 0)
    Assert.AreEqual "¤12345,6,80,0.00", Info.Format(1.234568E+08!, "c")
    
    Info.CurrencyGroupSizes = NewLongs()
    Assert.AreEqual "¤1234568000", Info.Format(1.234568E+09!, "c0")
End Sub

Public Sub TestPercent()
    Assert.AreEqual "0.00 %", Info.Format(0!, "p")
    Assert.AreEqual "123.00 %", Info.Format(1.23!, "p")
    Assert.AreEqual "123,456.00 %", Info.Format(1234.56!, "p")
    Assert.AreEqual "123,456,800.00 %", Info.Format(1234568!, "p")
    Assert.AreEqual "1,234,568,000.00 %", Info.Format(1.234568E+07!, "p")
    Assert.AreEqual "12,345 %", Info.Format(123.45!, "p0")
    Assert.AreEqual "12,345.68000000 %", Info.Format(123.4568!, "p8")
    Assert.AreEqual "-12,345 %", Info.Format(-123.45!, "p0")
    Assert.AreEqual "0.12 %", Info.Format(0.00123!, "p")
    Assert.AreEqual "0.124 %", Info.Format(0.001235!, "p3")
    Assert.AreEqual "12,300.00 %", Info.Format(123!, "p")
    Assert.AreEqual "12,312,310,000,000,000,000,000.00000 %", Info.Format(1.231231E+20!, "p5")
    
    Info.PercentNegativePattern = [-# %]
    Assert.AreEqual "-12,300.00 %", Info.Format(-123!, "p")

    Info.PercentNegativePattern = [-#%]
    Assert.AreEqual "-12,300.00%", Info.Format(-123!, "p")

    Info.PercentNegativePattern = [-%#]
    Assert.AreEqual "-%12,300.00", Info.Format(-123!, "p")

    Info.PercentGroupSizes = NewLongs(1, 2, 3)
    Assert.AreEqual "1,234,568,00,0.00 %", Info.Format(1.234568E+07!, "p")

    Info.PercentGroupSizes = NewLongs(1, 2, 1, 0)
    Assert.AreEqual "1234568,0,00,0.00 %", Info.Format(1.234568E+08!, "p")
    
    Info.PercentGroupSizes = cArray.CreateInstance(vbLong)
    Assert.AreEqual "123456800000 %", Info.Format(1.234568E+09!, "p0")
End Sub

Public Sub TestNAN()
    Dim d As Single
    On Error Resume Next
    d = 0! / 0!
    On Error GoTo 0
    Assert.AreEqual "NaN", Info.Format(d)
End Sub

Public Sub TestPosInfinity()
    Dim d As Single
    On Error Resume Next
    d = 1! / 0!
    On Error GoTo 0
    Assert.AreEqual "Infinity", Info.Format(d)
End Sub

Public Sub TestNegInfinity()
    Dim d As Single
    On Error Resume Next
    d = -1! / 0!
    On Error GoTo 0
    Assert.AreEqual "-Infinity", Info.Format(d)
End Sub


