VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArrayTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "Array Tests"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Dim LongArray()     As Long
    Dim Var             As Variant
    Dim ObjectArray()   As Object
        
    Select Case Test.MethodName
        Case "GetLength_WithValue_ReturnsExpected"
            ReDim LongArray(1 To 10)
            Test.Use(LongArray).Expect 10
            
            ReDim LongArray(1 To 10, 1 To 15)
            Test.Use(LongArray).Expect 150
        
            ReDim LongArray(1 To 10, 1 To 15)
            Test.Use(LongArray, 2).Expect 15
    
            ReDim Var(1 To 10) As Long
            Test.Use(Var).Expect 10
            
        Case "IsNull_WithValue_ReturnsExpected"
            Erase LongArray
            Test.Use(LongArray).Expect True
            
            Erase ObjectArray
            Test.Use(ObjectArray).Expect True
            
            Test.Use(Array()).Expect False
            
    End Select
End Sub

Public Function GetLength_WithValue_ReturnsExpected(ByRef Arr As Variant, Optional ByVal Dimension As Long) As Long
    GetLength_WithValue_ReturnsExpected = carray.GetLength(Arr, Dimension)
End Function

Public Sub GetLength_WithInvalidDimension_Throws()
    Dim Arr(1 To 10, 1 To 15)
    
    On Error GoTo errTrap
    carray.GetLength Arr, 3
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Dim Ex As Exception
    Catch Ex
    Assert.That TypeOf Ex Is IndexOutOfRangeException, Iz.True, "Incorrect exception was caught."
End Sub

Public Sub GetLength_WithNonArray_Throws()
    Dim NotAnArray As Long
    
    On Error GoTo errTrap
    carray.GetLength NotAnArray
    Assert.Fail "An exception should have been thrown."

errTrap:
    Dim Ex As Exception
    Catch Ex
    Assert.That TypeOf Ex Is ArgumentException, Iz.True, "Incorrect exception caught."
End Sub

Public Sub GetLength_WithUninitializedArray_Throws()
    Dim NullArray() As Long
    
    On Error GoTo errTrap
    carray.GetLength NullArray
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Dim Ex As Exception
    Catch Ex
    Assert.That TypeOf Ex Is ArgumentNullException, Iz.True, "Incorrect exception caught."
End Sub

Public Function IsNull_WithValue_ReturnsExpected(ByRef Arr As Variant) As Boolean
    IsNull_WithValue_ReturnsExpected = carray.IsNull(Arr)
End Function

Public Sub TestGetRankFor1DArray()
    Dim Arr(0) As Long
    Assert.AreEqual 1, carray.GetRank(Arr), "Rank should be 1."
End Sub

Public Sub TestGetRankFor2DArray()
    Dim Arr(0, 0) As Long
    Assert.AreEqual 2, carray.GetRank(Arr), "Rank should be 2."
End Sub

Public Sub TestGetRankFromVariantByRef()
    Dim Arr As Variant
    ReDim Arr(0, 0, 0) As Long
    Assert.AreEqual 3, carray.GetRank(Arr), "Rank should be 3."
End Sub

Public Sub TestGetRankNullArray()
    Dim Ex As Exception
    Dim NullArray() As Long
    
    On Error GoTo errTrap
    carray.GetRank NullArray
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
    carray.GetRank NotAnArray
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
    carray.Clear a, 0, 101
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
    carray.Clear a, 25, 30
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

    carray.Clear a, 10, 91
    
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
    
    carray.Clear a, 3, 5
    
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
    
    carray.ForEach a, AddressOf SetToNumber
    
    Dim i As Long
    For i = 0 To 10
        Assert.AreEqual 5, a(i)
    Next i
End Sub


