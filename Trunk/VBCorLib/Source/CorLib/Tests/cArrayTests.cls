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
            Test.Use(Array()).Expect 0
            
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
            
        Case "GetRank_WithValues_ReturnsExpected"
            ReDim LongArray(0)
            Test.Use(LongArray).Expect 1
            
            ReDim LongArray(0, 0)
            Test.Use(LongArray).Expect 2
            
            ReDim Var(0, 0, 0) As Long
            Test.Use(Var).Expect 3
    End Select
End Sub

Public Function GetLength_WithValue_ReturnsExpected(ByRef Arr As Variant, Optional ByVal Dimension As Long) As Long
    GetLength_WithValue_ReturnsExpected = cArray.GetLength(Arr, Dimension)
End Function

Public Sub GetLength_WithInvalidDimension_Throws()
    Dim Arr(1 To 10, 1 To 15)
    
    On Error Resume Next
    cArray.GetLength Arr, 3

    AssertIndexOutOfRangeException Err
End Sub

Public Sub GetLength_WithNonArray_Throws()
    Dim NotAnArray As Long
    
    On Error Resume Next
    cArray.GetLength NotAnArray

    AssertArgumentException Err
End Sub

Public Sub GetLength_WithUninitializedArray_Throws()
    Dim NullArray() As Long
    
    On Error Resume Next
    cArray.GetLength NullArray

    AssertArgumentNullException Err, "Arr"
End Sub

Public Function IsNull_WithValue_ReturnsExpected(ByRef Arr As Variant) As Boolean
    IsNull_WithValue_ReturnsExpected = cArray.IsNull(Arr)
End Function

Public Function GetRank_WithValues_ReturnsExpected(ByRef Arr As Variant) As Long
    GetRank_WithValues_ReturnsExpected = cArray.GetRank(Arr)
End Function

Public Sub GetRank_WithNullArray_Throws()
    Dim NullArray() As Long
    
    On Error Resume Next
    cArray.GetRank NullArray
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub GetRank_WithNonArray_Throws()
    Dim NotAnArray As Long

    On Error Resume Next
    cArray.GetRank NotAnArray
    
    AssertArgumentException Err, "Arr"
End Sub

Public Sub Clear_WithNullArray_Throws()
    Dim Arr() As Long
    
    On Error Resume Next
    cArray.Clear Arr, 0, 0
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub Clear_WithIndexLessThanLBound_ThrowsIndexOutOfRangeExceptions()
    Dim Arr(0) As Long
    
    On Error Resume Next
    cArray.Clear Arr, -1, 1
    
    AssertIndexOutOfRangeException Err
End Sub

Public Sub Clear_WithLengthLessThanZero_ThrowsIndexOutOfRangeException()
    Dim Arr(0) As Long
    
    On Error Resume Next
    cArray.Clear Arr, 0, -1
    
    AssertIndexOutOfRangeException Err
End Sub

Public Sub Clear_WithIndexAndLengthGreaterThanUBound_ThrowsIndexOutOfRangeException()
    Dim Arr(0) As Long
    
    On Error Resume Next
    cArray.Clear Arr, 0, 2
    
    AssertIndexOutOfRangeException Err
End Sub

Public Sub Clear_WithLongs_SetsAllToZero()
    Dim a() As Long
    a = NewLongs(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    
    cArray.Clear a, 0, 10
    
    Assert.That a, Has.All.EqualTo(0)
End Sub

Public Sub Clear_WithLongsSubSet_SetsSubSetToZero()
    Dim Actual() As Long
    Actual = NewLongs(1, 2, 3, 4, 5, 6)
    
    cArray.Clear Actual, 2, 2
    
    Dim Expected() As Long
    Expected = NewLongs(1, 2, 0, 0, 5, 6)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Clear_WithNonZeroBasedLongsSubSet_SetsSubsetToZero()
    Dim Actual(1 To 5) As Long
    Actual(1) = 1
    Actual(2) = 2
    Actual(3) = 3
    Actual(4) = 4
    Actual(5) = 5
    
    cArray.Clear Actual, 2, 2
    
    Dim Expected(1 To 5) As Long
    Expected(1) = 1
    Expected(2) = 0
    Expected(3) = 0
    Expected(4) = 4
    Expected(5) = 5
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Clear_WithUDTs_SetsAllToDefault()
    Dim a() As VBGUID
    ReDim a(0 To 1)
    a(0).Data1 = 1
    a(1).Data1 = 2

    cArray.Clear a, 0, 2
    
    Assert.That a(0).Data1, Iz.EqualTo(0)
    Assert.That a(1).Data1, Iz.EqualTo(0)
End Sub

Public Sub Clear_WithStrings_SetsToEmpty()
    Dim Actual() As String
    Actual = NewStrings("1", "2", "3", "4", "5")
    
    cArray.Clear Actual, 0, 5
    
    Assert.That Actual, Has.All.Empty
End Sub

Public Sub ForEach_WhenCalled_CallsMethodOnArrayElements()
    Dim Actual(10) As Long
    
    cArray.ForEach Actual, AddressOf SetToNumber
    
    Assert.That Actual, Has.All.EqualTo(5)
End Sub


