VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMixNumBinarySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private a() As Long

Public Sub Setup()
    ReDim a(-500 To 500)
    Dim i As Long
    For i = -500 To 500
        a(i) = i * 2
    Next i
End Sub

Public Sub TestFound()
    Assert.AreEqual -178, cArray.BinarySearch(a, -356&), "BinarySearch should return -355."
End Sub

Public Sub TestNotFound()
    Assert.AreEqual -998, cArray.BinarySearch(a, -7&), "BinarySearch should return -998."
End Sub

