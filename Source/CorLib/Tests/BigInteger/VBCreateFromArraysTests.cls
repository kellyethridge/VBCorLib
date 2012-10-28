VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBCreateFromArraysTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' VBCreateFromArraysTests
'
Option Explicit

Private b As BigInteger

Public Sub Setup()
    Set b = Nothing
End Sub



Public Sub NullArray()
    Dim v() As Byte
    Set b = NewBigInteger(v)
    AssertNumber Zero, 0, NewBytes()
End Sub

Public Sub EmptyArray()
    Set b = NewBigInteger(NewBytes())
    AssertNumber Zero, 0, NewBytes()
End Sub

Public Sub TestOne()
    Set b = NewBigInteger(NewBytes(1))
    AssertNumber Positive, 1, NewBytes(1)
End Sub

Public Sub TestOneWithLeadingZeros()
    Set b = NewBigInteger(NewBytes(1, 0, 0, 0, 0, 0, 0, 0, 0))
    AssertNumber Positive, 1, NewBytes(1)
End Sub

Public Sub TestNegativeOne()
    Set b = NewBigInteger(NewBytes(&HFF, &HFF))
    AssertNumber Negative, 2, NewBytes(&HFF, &HFF)
End Sub

Public Sub TestFF00000000()
    Set b = NewBigInteger(NewBytes(0, 0, 0, 0, &HFF))
    AssertNumber Positive, 5, NewBytes(0, 0, 0, 0, &HFF)
End Sub

Public Sub TestFFFF80000000()
    Set b = NewBigInteger(NewBytes(0, 0, 0, &H80, &HFF, &HFF))
    AssertNumber Negative, 4, NewBytes(0, 0, 0, &H80)
End Sub

Public Sub Test8000()
    Set b = NewBigInteger(NewBytes(0, &H80))
    AssertNumber Negative, 2, NewBytes(0, &H80)
End Sub

Public Sub TestFFFF()
    Set b = NewBigInteger(NewBytes(&HFF, &HFF))
    AssertNumber Negative, 2, NewBytes(&HFF, &HFF)
End Sub



Private Sub AssertNumber(ByVal Sign As Sign, ByVal Precision As Long, ByRef Arr() As Byte)
    Assert.IsNotNothing b, "b should not be Nothing"
    Assert.AreEqual Sign, b.Sign, "Wrong sign"
    Assert.AreEqual Precision, b.Precision, "Wrong precision"
    
    Dim a() As Byte
    a = b.ToArray
    Assert.AreEqual Arr, a, "Wrong array"
End Sub

