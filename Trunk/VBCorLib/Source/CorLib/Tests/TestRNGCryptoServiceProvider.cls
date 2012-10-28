VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRNGCryptoServiceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRNGCryptoServiceProvider
'
Option Explicit

Public Sub TestGetBytes()
    Const Size As Long = 1000
    
    Dim rng As New RNGCryptoServiceProvider
    Dim b1(1 To Size) As Byte
    Dim b2(1 To Size) As Byte
    rng.GetBytes b1
    rng.GetBytes b2
    
'    Dim eq As Long
'    Dim i As Long
'    For i = 1 To Size
'        If b1(i) = b2(i) Then
'            eq = eq + 1
'        End If
'    Next i
    
    Assert.That b1, Iz.Not.equalTo(b2), "b1 and b2 should contain random data."
End Sub

Public Sub TestGetNonZeroBytes()
    Const Size As Long = 1000
    
    Dim rng As New RNGCryptoServiceProvider
    Dim b1(1 To Size) As Byte
    Dim b2(1 To Size) As Byte
    rng.GetNonZeroBytes b1
    rng.GetNonZeroBytes b2
    
    Assert.That b1, Has.None.equalTo(0), "b1 should not contain zeros."
    Assert.That b2, Has.None.equalTo(0), "b2 should not contain zeros."
    Assert.That b1, Iz.Not.equalTo(b2), "b1 and b2 should contain random data."
    
'    Dim eq As Long
'    Dim i As Long
'    For i = 1 To Size
'        Assert.AreNotEqual 0, b1(i), "b1 should not contain a zero."
'        Assert.AreNotEqual 0, b2(i), "b2 should not contain a zero."
'        If b1(i) = b2(i) Then
'            eq = eq + 1
'        End If
'    Next i
'
'    Assert.That eq, Iz.Not.equalTo(Size), "b1 and b2 should contain random data."
End Sub
