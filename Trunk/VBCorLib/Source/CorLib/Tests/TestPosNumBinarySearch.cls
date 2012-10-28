VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestPosNumBinarySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private a() As Long

Public Sub Setup()
    ReDim a(0 To 1000)
    Dim i As Long
    For i = 0 To 1000
        a(i) = i * 2
    Next i
End Sub

Public Sub TestFound()
    Assert.AreEqual 3, cArray.BinarySearch(a, 6&), "BinarySearch should return 3."
End Sub

Public Sub TestNotFound()
    Assert.AreEqual -5, cArray.BinarySearch(a, 7&), "BinarySearch should return -5."
End Sub

Public Sub TestFoundInRange()
    Assert.AreEqual 3, cArray.BinarySearch(a, 6&, 1, 10), "Should return 3."
End Sub

Public Sub TestNotFoundInRange()
    Assert.AreEqual -11, cArray.BinarySearch(a, 6&, 10, 50), "Shoul return -11."
End Sub
