VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayIndexOf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestLong()
    Dim a() As Long
    Dim i As Long
    
    ReDim a(-1 To 100)
    For i = 0 To UBound(a)
        a(i) = i
    Next i
    
    i = cArray.IndexOf(a, CLng(UBound(a)))
    
    Assert.AreEqual UBound(a), i, "IndexOf should return " & UBound(a)
End Sub

Public Sub TestString()
    Dim a() As String
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To UBound(a)
        a(i) = "Item " & i
    Next i
    
    i = cArray.IndexOf(a, "Item " & UBound(a))
    Assert.AreEqual UBound(a), i, "IndexOf should return " & UBound(a)
End Sub

Public Sub TestDate()
    Dim a() As Date
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = DateSerial(i, 1, 1)
    Next i
    
    i = cArray.IndexOf(a, DateSerial(59, 1, 1))
    Assert.AreEqual 59, i, "IndexOf should return 59"
End Sub

Public Sub TestVariant()
    Dim a() As Variant
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = "Item " & i
    Next i
    Set a(88) = Me
    
    i = cArray.IndexOf(a, Me)
    Assert.AreEqual 88, i, "IndexOf should return 88"
End Sub

Public Sub TestDefaultComparer()
    Dim a() As String
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = "Item " & i
    Next i
    
    i = cArray.IndexOf(a, "Item 88", , , comparer.Default)
    Assert.AreEqual 88, i, "IndexOf should return 88."
End Sub

Public Sub TestNotInSubArray()
    Dim a() As Long
    Dim i As Long
    
    ReDim a(-1 To 100)
    For i = 0 To 100
        a(i) = i
    Next i
    
    i = cArray.IndexOf(a, 50&, 20, 20)
    Assert.AreEqual -2, i, "IndexOfEx should return -1."
End Sub

Public Sub TestObject()
    Dim a() As Int32
    Dim i As Long
    
    ReDim a(1 To 10)
    For i = 1 To 10
        Set a(i) = New Int32
        a(i).mValue = i
    Next i
    
    i = cArray.IndexOf(a, a(5))
    Assert.AreEqual 5, i
End Sub

Public Sub TestNothing()
    Dim a() As Variant
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = "Item " & i
    Next i
    Set a(UBound(a)) = Nothing
    i = cArray.IndexOf(a, Nothing)
    Assert.AreEqual UBound(a), i, "IndexOf should return " & UBound(a)
End Sub
