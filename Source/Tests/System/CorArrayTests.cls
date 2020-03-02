VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorArrayTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CorArrayTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Dim LongArray()     As Long
    Dim Var             As Variant
    Dim ObjectArray()   As Object
    Dim OneDimension(1 To 10)  As Long
    Dim TwoDimension(1 To 10, 1 To 5) As Long
    Dim NullArray() As Long
    
    Select Case Test.MethodName
        Case "Length_WithValue_ReturnsExpected"
            Test.Use Array(), 0
            Test.Use OneDimension, 10
            Test.Use TwoDimension, 50
            
        Case "IsNull_WithValue_ReturnsExpected"
            Test.Use NullArray, True
            Test.Use OneDimension, False
            Test.Use Array(), False
            Test.Use NullChars, True
            
        Case "IsNullOrEmpty_WithValue_ReturnsExpected"
            Test.Use NullArray, True
            Test.Use Array(), True
            Test.Use Array(1), False
            
        Case "Rank_WithValues_ReturnsExpected"
            ReDim LongArray(0)
            Test.Use LongArray, 1
            
            ReDim LongArray(0, 0)
            Test.Use LongArray, 2
            
            ReDim Var(0, 0, 0) As Long
            Test.Use Var, 3
            
        Case "GetLength_WithDimension_ReturnsExpected"
            Test.Use 1, 1
            Test.Use 2, 2
            Test.Use 3, 3
        
        Case "GetLength_WithInvalidDimensions_ThrowsIndexOutOfRangeException"
            Test.Use 0
            Test.Use 2
            
    End Select
End Sub

Public Sub GetLength_WithDimension_ReturnsExpected(ByVal Dimension As Long, ByVal Expected As Long)
    Dim Arr() As Long
    Dim Actual As Long
    
    ReDim Arr(0, 1, 2)
    
    Actual = CorArray.GetLength(Arr, Dimension)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetLength_WithInvalidDimensions_ThrowsIndexOutOfRangeException(ByVal InvalidDimension As Long)
    Dim Arr() As Long
    ReDim Arr(0)
    
    On Error Resume Next
    CorArray.GetLength Arr, InvalidDimension
    
    AssertIndexOutOfRangeException Err
End Sub

Public Sub GetLength_WithNonArray_ThrowsArgumentException()
    On Error Resume Next
    CorArray.GetLength 1, 1
    
    AssertArgumentException Err
End Sub

Public Sub Length_WithValue_ReturnsExpected(ByRef Arr As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = CorArray.Length(Arr)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Length_WithNonArray_ThrowsArgumentException()
    Dim NotAnArray As Long

    On Error Resume Next
    CorArray.Length NotAnArray

    AssertArgumentException Err
End Sub

Public Sub Length_WithUninitializedArray_Throws()
    Dim NullArray() As Long

    On Error Resume Next
    CorArray.Length NullArray

    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub IsNull_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = CorArray.IsNull(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsNull_WithNonArray_ThrowsArgumentException()
    On Error Resume Next
    
    CorArray.IsNull "abc"
    
    AssertArgumentException Err, "Arr"
End Sub

Public Sub IsNullOrEmpty_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = CorArray.IsNullOrEmpty(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsNullOrEmpty_WithNonArray_ThrowsArgumentException()
    On Error Resume Next
    
    CorArray.IsNullOrEmpty "abc"
    
    AssertArgumentException Err, "Arr"
End Sub

Public Sub Rank_WithValues_ReturnsExpected(ByRef Arr As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = CorArray.Rank(Arr)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Rank_WithNullArray_ThrowsArgumentNullException()
    Dim NullArray() As Long

    On Error Resume Next
    CorArray.Rank NullArray

    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub Rank_WithNonArray_ThrowsArgumentException()
    Dim NotAnArray As Long

    On Error Resume Next
    CorArray.Rank NotAnArray

    AssertArgumentException Err
End Sub

Public Sub Clear_WithNullArray_Throws()
    Dim Arr() As Long

    On Error Resume Next
    CorArray.Clear Arr, 0, 0

    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub Clear_WithIndexLessThanLBound_ThrowsIndexOutOfRangeExceptions()
    Dim Arr(0) As Long

    On Error Resume Next
    CorArray.Clear Arr, -1, 1

    AssertIndexOutOfRangeException Err
End Sub

Public Sub Clear_WithLengthLessThanZero_ThrowsIndexOutOfRangeException()
    Dim Arr(0) As Long

    On Error Resume Next
    CorArray.Clear Arr, 0, -1

    AssertIndexOutOfRangeException Err
End Sub

Public Sub Clear_WithIndexAndLengthGreaterThanUBound_ThrowsIndexOutOfRangeException()
    Dim Arr(0) As Long

    On Error Resume Next
    CorArray.Clear Arr, 0, 2

    AssertIndexOutOfRangeException Err
End Sub

Public Sub Clear_WithLongs_SetsAllToZero()
    Dim a() As Long
    a = NewLongs(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    CorArray.Clear a, 0, 10

    Assert.That a, Has.All.EqualTo(0)
End Sub

Public Sub Clear_WithLongsSubSet_SetsSubSetToZero()
    Dim Actual() As Long
    Actual = NewLongs(1, 2, 3, 4, 5, 6)

    CorArray.Clear Actual, 2, 2

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

    CorArray.Clear Actual, 2, 2

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

    CorArray.Clear a, 0, 2

    Assert.That a(0).Data1, Iz.EqualTo(0)
    Assert.That a(1).Data1, Iz.EqualTo(0)
End Sub

Public Sub Clear_WithStrings_SetsToEmpty()
    Dim Actual() As String
    Actual = NewStrings("1", "2", "3", "4", "5")

    CorArray.Clear Actual, 0, 5

    Assert.That Actual, Has.All.Empty
End Sub

Public Sub ForEach_WhenCalled_CallsMethodOnArrayElements()
    Dim Actual(10) As Long

    CorArray.ForEach Actual, AddressOf SetToNumber

    Assert.That Actual, Has.All.EqualTo(5)
End Sub


