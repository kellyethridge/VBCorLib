VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMathExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMathExt
'
Option Explicit

Public Sub TestCeiling()
    Assert.AreEqual 3#, Ceiling(2.9)
    Assert.AreEqual -2#, Ceiling(-2.9)
End Sub

Public Sub TestFloor()
    Assert.AreEqual 2#, Floor(2.9)
    Assert.AreEqual -3#, Floor(-2.9)
End Sub

Public Sub TestAcos()
    Assert.AreEqualFloats 1.33403041100045, ACos(0.23456), 0.000000001
End Sub

Public Sub TestMax()
    Assert.AreEqual 5, Max(2, 5)
    Assert.AreEqual 5, Max(5, 2)
End Sub

Public Sub TestMaxObject()
    Dim X As New Int32
    Dim Y As New Int32
    
    X.Value = 2
    Y.Value = 5
    
    Assert.AreEqual Y, Max(X, Y)
    Assert.AreEqual Y, Max(Y, X)
End Sub

Public Sub TestMin()
    Assert.AreEqual 2, Min(2, 5)
    Assert.AreEqual 2, Min(5, 2)
End Sub

Public Sub TestMinObject()
    Dim X As New Int32
    Dim Y As New Int32
    
    X.Value = 2
    Y.Value = 5
    
    Assert.AreEqual X, Min(X, Y)
    Assert.AreEqual X, Min(Y, X)
End Sub

Public Sub TestDivRem()
    Dim Q As Long
    Dim r As Long
    
    Q = DivRem(11, 4, r)
    
    Assert.AreEqual 2, Q
    Assert.AreEqual 3, r
End Sub

Public Sub TestDivRemNegNum()
    Dim Q As Long
    Dim r As Long
    
    Q = DivRem(-11, 4, r)
    
    Assert.AreEqual -2, Q
    Assert.AreEqual -3, r
End Sub

Public Sub TestDivRemNegNumAndDenom()
    Dim Q As Long
    Dim r As Long
    
    Q = DivRem(-11, -4, r)
    
    Assert.AreEqual 2, Q
    Assert.AreEqual -3, r
End Sub

Public Sub TestDivRemNegDenom()
    Dim Q As Long
    Dim r As Long
    
    Q = DivRem(11, -4, r)
    
    Assert.AreEqual -2, Q
    Assert.AreEqual 3, r
End Sub

Public Sub TestMinInArray()
    Dim a(10)
    Dim i As Long
    For i = 0 To 10
        a(i) = 10 - i
    Next i
    
    Assert.AreEqual 0, MinIn(a)
End Sub

Public Sub TestMaxInArray()
    Dim a(10)
    Dim i As Long
    For i = 0 To 10
        a(i) = i
    Next i
    
    Assert.AreEqual 10, MaxIn(a)
End Sub
