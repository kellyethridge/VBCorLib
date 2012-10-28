VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayReverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestAllLongs()
    Dim a(100) As Long
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i
    Next i
    cArray.Reverse a
    For i = 0 To 100
        Assert.AreEqual 100 - i, a(i)
    Next i
End Sub

Public Sub TestAllStrings()
    Dim a(100) As String
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i
    Next i
    cArray.Reverse a
    For i = 0 To 100
        Assert.AreEqual CStr(100 - i), a(i)
    Next i
End Sub

Public Sub TestSubArrayLongs()
    Dim a(100) As Long
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i
    Next i
    cArray.Reverse a, 26, 24
    For i = 0 To 100
        If i <= 25 Or i >= 50 Then
            Assert.AreEqual i, a(i), "Reversed outside subarray."
        Else
            Assert.AreEqual 100 - i - 25, a(i), "Subarray not reversed."
        End If
    Next i
End Sub

Public Sub TestLengthTooLong()
    Dim a(100)
    
    On Error GoTo errTrap
    cArray.Reverse a, 0, 102
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestUDTs()
    Dim a(10) As VBGUID
    Dim i As Long
    
    For i = 0 To 10
        a(i).Data1 = i
    Next i
    
    cArray.Reverse a
    
    For i = 0 To 10
        Assert.AreEqual 10 - i, a(i).Data1
    Next i
    
End Sub
