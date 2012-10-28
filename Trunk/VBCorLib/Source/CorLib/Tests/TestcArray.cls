VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestcArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestLengthOf1DArray()
    Dim arr(1 To 10) As Long
    Assert.AreEqual 10, cArray.GetLength(arr), "Length should be 10."
End Sub

Public Sub TestLengthof2DArray()
    Dim arr(1 To 10, 1 To 15) As Long
    Assert.AreEqual 150, cArray.GetLength(arr), "Length should be 100."
End Sub

Public Sub TestSpecificDimension()
    Dim arr(1 To 10, 1 To 15)
    Assert.AreEqual 15, cArray.GetLength(arr, 2)
End Sub

Public Sub TestGetLengthInVariantByRef()
    Dim arr As Variant
    ReDim arr(1 To 10) As Long
    Assert.AreEqual 10, cArray.GetLength(arr), "Lenght should be 10."
End Sub

Public Sub TestGetLengthRankError()
    Dim arr(1 To 10, 1 To 15)
    On Error GoTo errTrap
    cArray.GetLength arr, 3
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is IndexOutOfRangeException, "Incorrect exception was caught."
End Sub

Public Sub TestIsArrayError()
    Dim NotAnArray As Long
    Dim Ex As Exception
    
    On Error GoTo errTrap
    cArray.GetLength NotAnArray
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."

errTrap:
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Incorrect exception caught."
End Sub

Public Sub TestGetLengthNullArray()
    Dim Ex As Exception
    Dim NullArray() As Long
    
    On Error GoTo errTrap
    cArray.GetLength NullArray
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArgumentNullException, "Incorrect exception caught."
End Sub

Public Sub TestIsNullArrayTrue()
    Dim NullArray() As Long
    Assert.IsTrue cArray.IsNull(NullArray), "IsNull should return True for 'NullArray'."
    
    Dim NullObjArray() As Object
    Assert.IsTrue cArray.IsNull(NullObjArray), "IsNull should return True for 'NullObjArray'."
    
    ReDim NullArray(0)
    Assert.IsFalse cArray.IsNull(NullArray), "IsNull should return false for 'NullArray'."
End Sub

Public Sub TestGetRankFor1DArray()
    Dim arr(0) As Long
    Assert.AreEqual 1, cArray.GetRank(arr), "Rank should be 1."
End Sub

Public Sub TestGetRankFor2DArray()
    Dim arr(0, 0) As Long
    Assert.AreEqual 2, cArray.GetRank(arr), "Rank should be 2."
End Sub

Public Sub TestGetRankFromVariantByRef()
    Dim arr As Variant
    ReDim arr(0, 0, 0) As Long
    Assert.AreEqual 3, cArray.GetRank(arr), "Rank should be 3."
End Sub

Public Sub TestGetRankNullArray()
    Dim Ex As Exception
    Dim NullArray() As Long
    
    On Error GoTo errTrap
    cArray.GetRank NullArray
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArgumentNullException, "Incorrect exception caught."
End Sub

Public Sub TestGetRankNotAnArray()
    Dim NotAnArray As Long
    Dim Ex As Exception
    On Error GoTo errTrap
    cArray.GetRank NotAnArray
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Incorrect exception caught."
End Sub

Public Sub TestClearAllLongs()
    Dim a() As Long
    ReDim a(100)
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i
    Next i
    cArray.Clear a, 0, 101
    For i = 0 To 100
        Assert.AreEqual 0, a(i)
    Next i
End Sub

Public Sub TestClearSomeLongs()
    Dim a() As Long
    ReDim a(100)
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i
    Next i
    cArray.Clear a, 25, 30
    For i = 0 To 100
        If i < 25 Or i > 54 Then
            Assert.AreEqual i, a(i)
        Else
            Assert.AreEqual 0, a(i)
        End If
    Next i
End Sub

Public Sub TestClearSomeUDTs()
    Dim a() As VBGUID
    ReDim a(10 To 100)
    Dim i As Long

    For i = 10 To 100
        a(i).Data1 = i
    Next i

    cArray.Clear a, 10, 91
    
    For i = 10 To 100
        Assert.AreEqual 0, a(i).Data1
    Next i
End Sub

Public Sub TestClearSomeStrings()
    Dim a() As String
    ReDim a(10)
    Dim i As Long
    
    For i = 0 To 10
        a(i) = "Item " & i
    Next i
    
    cArray.Clear a, 3, 5
    
    For i = 0 To 10
        If i < 3 Or i > 7 Then
            Assert.AreEqual "Item " & i, a(i)
        Else
            Assert.AreEqual "", a(i)
        End If
    Next i
End Sub

Public Sub ForEachAssignNumber()
    Dim a(10) As Long
    
    cArray.ForEach a, AddressOf SetToNumber
    
    Dim i As Long
    For i = 0 To 10
        Assert.AreEqual 5, a(i)
    Next i
End Sub
