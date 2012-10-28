VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArraySort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestSortLongs()
    Dim a() As Long
    FillArray a, 100, 1000
    cArray.Sort a
    AssertValueTypes a
End Sub

Public Sub TestSortStrings()
    Dim a() As String
    FillArray a, 100, 1000
    cArray.Sort a
    AssertValueTypes a
End Sub

Public Sub TestSortObjects()
    Dim a() As Int32
    Dim i As Long
    
    ReDim a(0 To 100)
    Rnd -13
    For i = 0 To UBound(a)
        Set a(i) = New Int32
        a(i).mValue = Int(Rnd * 100)
    Next i
    
    cArray.Sort a
    
    For i = 0 To UBound(a) - 1
        Assert.IsTrue a(i).mValue <= a(i + 1).mValue, "Elements out of order at " & i & " and " & i + 1 & "."
    Next i
    
End Sub

Public Sub TestSortInteger()
    Dim a() As Integer
    FillArray a, 100, 1000
    cArray.Sort a
    AssertValueTypes a
End Sub

Public Sub TestSortBytes()
    Dim a() As Byte
    FillArray a, 100, 255
    cArray.Sort a
    AssertValueTypes a
End Sub

Public Sub TestSortKeysAndItems()
    Dim k() As Long
    Dim i() As Byte

    FillArray k, 100, 1000
    FillArray i, 100, 255
    cArray.SortKey k, i
    AssertValueTypes k
    AssertValueTypes i
End Sub

Public Sub TestWithComparer()
    Dim k() As Long
    Dim i() As Byte

    FillArray k, 100, 1000
    FillArray i, 100, 255
    cArray.SortKey k, i, comparer.Default
    AssertValueTypes k
    AssertValueTypes i
End Sub

Public Sub TestWithStringComparer()
    Dim s(100) As String
    Dim i As Long
    
    Rnd -13
    For i = 0 To 100
        s(i) = "Value " & Chr$(Int(Rnd * 255))
    Next i
    cArray.Sort s, comparer.Default
    AssertValueTypes s
End Sub

Public Sub TestSortVariants()
    Dim a() As Variant
    FillArray a, 100, 1000
    cArray.Sort a
    AssertValueTypes a
End Sub

Public Sub TestGeneralSort()
    Dim a() As Long
    FillArray a, 100, 1000
    cArray.Sort a, comparer.Default
    AssertValueTypes a
End Sub

Public Sub TestNonComparableElements()
    Dim a() As Variant
    Dim b() As Long
    FillArray a, 100, 1000
    a(39) = b
    On Error GoTo errTrap
    cArray.Sort a
    On Error GoTo 0
    Assert.Fail "Should have thrown an error."
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "Should have caught an exception."
    Assert.IsTrue TypeOf Ex Is InvalidOperationException, "Wrong exception type caught."
End Sub

Public Sub TestLengthTooLong()
    Dim a(10) As Long
    
    On Error GoTo errTrap
    cArray.BinarySearch a, 1, 0, 12
    On Error GoTo 0
    Assert.Fail "An exception should have been raised."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestSortSubArray()
    Dim a() As Long
    FillArray a, 100, 1000
    cArray.SortEx a, 26, 24
    
    Dim i As Long
    Dim j As Long
    Rnd -13
    For i = 0 To 98
        j = Int(Rnd * 1000)
        If i <= 25 Or i >= 50 Then
            Assert.AreEqual j, a(i), "Sorted outside sub-array."
        Else
            ' we don't want to bleed into the unsorted list with a(i + 1).
            If i < 49 Then Assert.IsTrue a(i) <= a(i + 1), "sub-array not sorted."
        End If
    Next i
End Sub

Public Sub TestSortStringCallback()
    Dim a(100) As String
    Dim i As Long
    
    Rnd -13
    For i = 0 To 100
        a(i) = "String " & Int(Rnd * 1000)
    Next i
    
    cArray.Sort a, AddressOf CompareStrings
    
    AssertValueTypes a
End Sub

Public Sub TestSortUDTCallback()
    Dim a(100) As VBGUID
    Dim l(100) As Long
    Dim i As Long
    
    Rnd -13
    For i = 0 To 100
        a(i).Data1 = Rnd * 1000
        l(i) = a(i).Data1
    Next i
    
    cArray.Sort a, AddressOf CompareVBGuids
    cArray.Sort l
    
    For i = 0 To 100
        Assert.AreEqual l(i), a(i).Data1
    Next i
End Sub

Public Sub TestNonZeroLBound()
    Dim a(1 To 10) As Double
    Dim i As Long
    
    Rnd -13
    For i = 1 To 10
        a(i) = Rnd * 10000
    Next i
    
    cArray.Sort a
    
    AssertValueTypes a
End Sub


Private Sub FillArray(ByRef arr As Variant, ByVal Size As Long, ByVal maxValue As Long)
    Dim i As Long

    ReDim arr(0 To Size - 1)
    Rnd -13
    For i = LBound(arr) To UBound(arr)
        arr(i) = Int(Rnd * maxValue)
    Next i
End Sub

Private Sub AssertValueTypes(ByRef arr As Variant)
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr) - 1
        Assert.IsTrue arr(i) <= arr(i + 1), "Element out of order at " & i & "."
    Next i
End Sub


