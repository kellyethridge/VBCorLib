VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorArrayBinarySearchTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "BinarySearch_WithValues_ReturnsExpected"
            Test.Use NewBytes(1, 2, 3, 4, 5), 1, 0
            Test.Use NewBytes(1, 2, 3, 4, 5), 3, 2
            Test.Use NewBytes(1, 2, 3, 4, 5), 5, 4
            Test.Use NewBytes(2, 3, 4, 5), 1, -1
            Test.Use NewBytes(1, 2, 4, 5), 3, -3
            Test.Use NewBytes(1, 2, 3, 4), 5, -5
            Test.Use NewIntegers(1, 2, 3, 4, 5), 1, 0
            Test.Use NewIntegers(1, 2, 3, 4, 5), 3, 2
            Test.Use NewIntegers(1, 2, 3, 4, 5), 5, 4
            Test.Use NewLongs(1, 2, 3, 4, 5), 1, 0
            Test.Use NewLongs(1, 2, 3, 4, 5), 3, 2
            Test.Use NewLongs(1, 2, 3, 4, 5), 5, 4
            Test.Use NewLongs(1), 1, 0
            Test.Use NewLongs(1), 0, -1
            Test.Use NewLongs(1), 2, -2
            Test.Use NewDoubles(1, 2, 3, 4, 5), 1, 0
            Test.Use NewDoubles(1, 2, 3, 4, 5), 3, 2
            Test.Use NewDoubles(1, 2, 3, 4, 5), 5, 4
            Test.Use NewSingles(1, 2, 3, 4, 5), 1, 0
            Test.Use NewSingles(1, 2, 3, 4, 5), 3, 2
            Test.Use NewSingles(1, 2, 3, 4, 5), 5, 4
            Test.Use NewCurrencys(1, 2, 3, 4, 5), 1, 0
            Test.Use NewCurrencys(1, 2, 3, 4, 5), 3, 2
            Test.Use NewCurrencys(1, 2, 3, 4, 5), 5, 4
            Test.Use NewDates(#1/1/2001#, #1/2/2001#, #1/3/2001#, #1/4/2001#, #1/5/2001#), #1/1/2001#, 0
            Test.Use NewDates(#1/1/2001#, #1/2/2001#, #1/3/2001#, #1/4/2001#, #1/5/2001#), #1/3/2001#, 2
            Test.Use NewDates(#1/1/2001#, #1/2/2001#, #1/3/2001#, #1/4/2001#, #1/5/2001#), #1/5/2001#, 4
            Test.Use NewStrings("a", "b", "c", "d", "e"), "a", 0
            Test.Use NewStrings("a", "b", "c", "d", "e"), "c", 2
            Test.Use NewStrings("a", "b", "c", "d", "e"), "e", 4
            Test.Use NewVariants(CDec(1), CDec(2), CDec(3), CDec(4), CDec(5)), CDec(1), 0
            Test.Use NewVariants(CDec(1), CDec(2), CDec(3), CDec(4), CDec(5)), CDec(3), 2
            Test.Use NewVariants(CDec(1), CDec(2), CDec(3), CDec(4), CDec(5)), CDec(5), 4
            Test.Use Array(MakeInt32(1), MakeInt32(2), MakeInt32(3), MakeInt32(4), MakeInt32(5)), MakeInt32(1), 0
            Test.Use Array(MakeInt32(1), MakeInt32(2), MakeInt32(3), MakeInt32(4), MakeInt32(5)), MakeInt32(3), 2
            Test.Use Array(MakeInt32(1), MakeInt32(2), MakeInt32(3), MakeInt32(4), MakeInt32(5)), MakeInt32(5), 4
            
    End Select
End Sub

Public Sub BinarySearch_WithValues_ReturnsExpected(ByRef Arr As Variant, ByRef Value As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = CorArray.BinarySearch(Arr, Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BinarySearchEx_WithIndexAndLength_SearchesSubArray()
    Dim Arr() As Long
    Arr = NewLongs(1, 2, 3, 4, 5)
    Dim Actual As Long
    
    Actual = CorArray.BinarySearchEx(Arr, 2, 3, 3)
    
    Const Expected As Long = 2
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BinarySearchEx_WithIndexAndLength_DoesNotFindInSubArray()
    Dim Arr() As Long
    Arr = NewLongs(1, 2, 3, 4, 5)
    Dim Actual As Long
    
    Actual = CorArray.BinarySearchEx(Arr, 0, 3, 4)
    
    Const Expected As Long = -4
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BinarySearch_WithNullArray_Throws()
    Dim Arr() As Long
    
    On Error Resume Next
    CorArray.BinarySearch Arr, 1
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub BinarySearchEx_WithNullArray_Throws()
    Dim Arr() As Long
    
    On Error Resume Next
    CorArray.BinarySearchEx Arr, 0, 1, 0
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub BinarySearchEx_WithIndexLessThanLBound_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.BinarySearchEx Arr, -1, 1, 0
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub BinarySearchEx_WithLengthLessThanZero_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.BinarySearchEx Arr, 0, -1, 0
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub TestLengthTooLong()
    Dim a(10) As Long
    
    On Error GoTo errTrap
    CorArray.BinarySearchEx a, 0, 12, 1
    On Error GoTo 0
    Assert.Fail "An exception should have been raised."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should have been caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
End Sub


