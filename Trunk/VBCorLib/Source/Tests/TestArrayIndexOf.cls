VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArrayIndexOfTests"
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
        Case "IndexOf_WithValues_ReturnsExpected"
            Test.Use NewBytes(1, 2, 3, 4, 5, 6), 4, 3
            Test.Use NewLongs(1, 2, 3, 4, 5, 6), 7, -1
            Test.Use NewIntegers(1, 2, 3, 4, 5, 6), 4, 3
            Test.Use NewLongs(1, 2, 3, 4, 5, 6), 4, 3
            Test.Use NewDoubles(1, 2, 3, 4, 5, 6), 4, 3
            Test.Use NewSingles(1, 2, 3, 4, 5, 6), 4, 3
            Test.Use NewCurrencys(1, 2, 3, 4, 5, 6), 4, 3
            Test.Use NewStrings("1", "2", "3", "4", "5", "6"), "4", 3
            Test.Use NewDates(#1/1/2001#, #2/1/2001#, #3/1/2001#, #4/1/2001#, #5/1/2001#, #6/1/2001#), #4/1/2001#, 3
            Test.Use NewVariants("1", "2", "3", "4", "5", "6"), "4", 3
            
            Dim Arr(2) As Int32
            Set Arr(0) = New Int32
            Set Arr(1) = New Int32
            Set Arr(2) = New Int32
            Test.Use Arr, Arr(2), 2
            Test.Use Arr, New Int32, -1
            Set Arr(2) = Nothing
            Test.Use Arr, Nothing, 2
            
            Dim ArrLongs(1 To 3) As Long
            Test.Use ArrLongs, 1, 0
            
    End Select
End Sub

Public Sub IndexOf_WithValues_ReturnsExpected(ByRef Arr As Variant, ByRef Value As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = cArray.IndexOf(Arr, Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IndexOf_WithNullArray_ThrowsArgumentNullException()
    Dim Arr() As Long
    
    On Error GoTo errTrap
    cArray.IndexOf Arr, 1
    
errTrap:
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub IndexOf_WithMuliDimensionArray_ThrowsRankException()
    Dim Arr(0, 0) As Long
    
    On Error GoTo errTrap
    cArray.IndexOf Arr, 0
    
errTrap:
    AssertRankException Err
End Sub

Public Sub IndexOf_WithStartIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Arr(0) As Long
    
    On Error GoTo errTrap
    cArray.IndexOf Arr, 0, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub IndexOf_WithStartIndexGreaterThanUBound_ThrowsArgumentOutOfRangeException()
    Dim Arr(0) As Long
    
    On Error GoTo errTrap
    cArray.IndexOf Arr, 0, 1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub IndexOf_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Arr(0) As Long
    
    On Error GoTo errTrap
    cArray.IndexOf Arr, 0, 0, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub IndexOf_WithStartIndexAndCountGreaterThanUBound_ThrowsArgumentOutOfRangeException()
    Dim Arr(0) As Long
    
    On Error GoTo errTrap
    cArray.IndexOf Arr, 0, 0, 2
    
errTrap:
    AssertArgumentOutOfRangeException Err
End Sub
