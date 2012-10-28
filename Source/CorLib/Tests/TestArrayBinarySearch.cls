VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayBinarySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestSearchSubArray()
    Dim a() As Long
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        If i >= 26 And i <= 49 Then a(i) = i
    Next i
    
    For i = 26 To 49
        Assert.AreEqual i, cArray.BinarySearch(a, i, 26, 24)
    Next i
End Sub

Public Sub TestNotInSubArray()
    Dim a() As Long
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        If i >= 26 And i <= 49 Then a(i) = i
    Next i
    
    For i = 0 To 25
        Assert.AreEqual Not 26, cArray.BinarySearch(a, i, 26, 24, comparer.Default)
    Next i
    For i = 50 To 100
        Assert.AreEqual Not 50, cArray.BinarySearch(a, i, 26, 24)
    Next i
End Sub

Public Sub TestObjectBinarySearch()
    Dim a() As Int32
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        Set a(i) = New Int32
        a(i).mValue = i
    Next i
    Dim oInt As New Int32
    oInt.mValue = 5
    i = cArray.BinarySearch(a, oInt)
    Assert.AreEqual 5, i, "Should return 5."
End Sub

Public Sub TestStringBinarySearch()
    Dim a() As String
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = "Item " & Chr$(i)
    Next i
    
    i = cArray.BinarySearch(a, "Item " & Chr$(5))
    Assert.AreEqual 5, i, "Should return 5."
End Sub

Public Sub TestVariantBinarySearch()
    Dim a() As Variant
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = "Item " & Chr$(i)
    Next i
    a(0) = Null
    a(1) = Empty
    Assert.AreEqual 0, cArray.BinarySearch(a, Null), "Should return 0."
    Assert.AreEqual 1, cArray.BinarySearch(a, Empty), "Should return 1."
End Sub

Public Sub TestLengthTooLong()
    Dim a(10) As Long
    
    On Error GoTo errTrap
    cArray.BinarySearch a, 1, 0, 12
    On Error GoTo 0
    Assert.Fail "An exception should have been raised."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
End Sub

