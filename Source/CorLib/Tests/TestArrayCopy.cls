VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestArrayCopy
'
Option Explicit

Public Sub TestCopyLongs()
    Dim src() As Long
    Dim dst() As Long
    Dim i As Long
    
    ReDim src(0 To 10)
    ReDim dst(0 To UBound(src))
    For i = 0 To UBound(src)
        src(i) = i
    Next i
    
    carray.Copy src, dst, UBound(src) + 1
    
    Assert.AreEqual src, dst
End Sub

Public Sub TestCopyStrings()
    Dim src() As String
    Dim dst() As String
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To UBound(src))
    For i = 0 To UBound(src)
        src(i) = CStr(i)
    Next i
    
    carray.Copy src, dst, UBound(src) + 1
    
    Assert.AreEqual src, dst
End Sub

Public Sub TestCopySubArray()
    Dim src() As Long
    Dim dst() As Long
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To 100)
    For i = 0 To 100
        src(i) = i
    Next i
    
    carray.CopyEx src, 20, dst, 0, 20
    For i = 0 To 100
        If i >= 20 Then
            Assert.AreEqual 0, dst(i)
        Else
            Assert.AreEqual i + 20, dst(i)
        End If
    Next i

End Sub

Public Sub TestCopyObjects()
    Dim src() As Int32
    Dim dst() As Int32
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To 100)
    For i = 0 To 100
        Set src(i) = New Int32
        src(i).mValue = i
    Next i
    
    carray.Copy src, dst, 101
    For i = 0 To 100
        Assert.AreEqual i, dst(i).mValue
    Next i
End Sub

Public Sub TestObjectsToSelf()
    Dim Item1 As New Int32
    Dim Item2 As New Int32
    Dim Item3 As New Int32
    Dim Actual(0 To 3) As Int32
    Set Actual(0) = Item1
    Set Actual(1) = Item2
    Set Actual(2) = Item3
    Set Actual(3) = New Int32
    
    carray.CopyEx Actual, 0, Actual, 1, 3
    
    Dim Expected(0 To 3) As Int32
    Set Expected(0) = Item1
    Set Expected(1) = Item1
    Set Expected(2) = Item2
    Set Expected(3) = Item3
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TestLongToString()
    Dim src() As Long
    Dim dst() As String
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To 100)
    For i = 0 To 100
        src(i) = i
    Next i
    
    carray.Copy src, dst, 101
    For i = 0 To 100
        Assert.AreEqual CStr(i), dst(i)
    Next i
End Sub

Public Sub TestIndexes()
    Dim src() As String
    Dim dst() As String
    
    ReDim src(0 To 4)
    ReDim dst(5 To 15)
    
    Dim i As Long
    For i = 0 To 4
        src(i) = i
    Next i
    
    carray.CopyEx src, 0, dst, 5, 5
    
    For i = 5 To 9
        Assert.AreEqual CStr(i - 5), dst(i)
    Next i
End Sub

Public Sub TestVariantToObjects()
    Dim src() As Variant
    Dim dst() As Int32
    
    ReDim src(0 To 5)
    ReDim dst(0 To 5)
    Dim i As Long
    For i = 0 To 5
        Set src(i) = New Int32
        Set dst(i) = New Int32
    Next i
    
    carray.Copy src, dst, 6
    
    For i = 0 To 5
        Assert.AreEqual src(i), dst(i)
    Next i
End Sub

Public Sub TestArrayTypeMismatch()
    Dim src(0) As Double
    Dim dst(0) As Long
    
    On Error GoTo errTrap
    carray.Copy src, dst, 1
    On Error GoTo 0
    Assert.Fail "Error should have been raised."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArrayTypeMismatchException, "Wrong exception type caught."
    
End Sub

Public Sub TestCopyToSelf()
    Dim a(100) As Long
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i + 1
    Next i
    
    carray.CopyEx a, 0, a, 1, 100
    
    Assert.AreEqual 1, a(0)
    
    For i = 1 To 100
        Assert.AreEqual i, a(i)
    Next i
End Sub

Public Sub TestUserDefined()
    Dim src(10) As VBGUID
    Dim dst(10) As VBGUID
    Dim i As Long

    For i = 0 To 10
        src(i).Data1 = i
    Next i

    carray.Copy src, dst, 11

    For i = 0 To 10
        Assert.AreEqual i, dst(i).Data1
    Next i

End Sub

Public Sub TestCopyObjectsWithCast()
    Dim src(10) As Int32
    Dim dst(10) As IComparable
    Dim i As Long
    
    For i = 0 To 10
        Set src(i) = New Int32
        src(i).mValue = i
    Next i
    
    carray.Copy src, dst, 11
    
    For i = 0 To 10
        Assert.IsNotNothing dst(i)
        Assert.AreEqual 0, dst(i).CompareTo(src(i))
    Next i
    
End Sub

Public Sub TestCopyObjectsMismatch()
    Dim src(10) As Int32
    Dim dst(10) As Int32Formattable
    Dim i As Long
    
    For i = 0 To 10
        Set src(10) = New Int32
    Next i
    
    On Error GoTo errTrap
    carray.Copy src, dst, 11
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArrayTypeMismatchException, "Wrong exception type caught."
End Sub

Public Sub TestNonZeroLBound()
    Dim src(1 To 10) As String
    Dim dst(1 To 10) As String
    Dim i As Long
    
    For i = 1 To 10
        src(i) = "Item" & i
    Next i
    
    carray.Copy src, dst, 10
    
    For i = 1 To 10
        Assert.AreEqual "Item" & i, dst(i)
    Next i
End Sub
