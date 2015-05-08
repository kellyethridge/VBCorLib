VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArrayFindTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCorArrayFinds
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Dim Arr(-1 To 1) As Long
    Dim ArrInt32() As Int32
    
    Select Case Test.MethodName
        Case "FindIndex_WithValues_ReturnsExpected"
            Test.Use NewLongs(1, 2, 3, 3, 5), 3, 2
            Test.Use NewLongs(1, 2, 3, 3, 5), 6, -1
            Test.Use Arr, 1, -2
        
        Case "Find_WithValues_ReturnsExpected"
            ReDim ArrInt32(1)
            Set ArrInt32(0) = MakeInt32(1)
            Set ArrInt32(1) = MakeInt32(2)
            Test.Use NewBytes(1, 2, 3, 4, 5), AddressOf FindByteCallback, 3, CByte(3)
            Test.Use NewIntegers(1, 2, 3, 4, 5), AddressOf FindIntegerCallback, 3, 3
            Test.Use NewLongs(1, 2, 3, 4, 5), AddressOf FindLongCallback, 3, 3&
            Test.Use NewSingles(1, 2, 3, 4, 5), AddressOf FindSingleCallback, 3, 3!
            Test.Use NewDoubles(1, 2, 3, 4, 5), AddressOf FindDoubleCallback, 3, 3#
            Test.Use NewCurrencys(1, 2, 3, 4, 5), AddressOf FindCurrencyCallback, 3, 3@
            Test.Use ArrInt32, AddressOf FindInt32Callback, 2, ArrInt32(1)
            
            Test.Use NewBytes(1, 2, 3, 4, 5), AddressOf FindByteCallback, 0, CByte(0)
            Test.Use NewIntegers(1, 2, 3, 4, 5), AddressOf FindIntegerCallback, 0, 0
            Test.Use NewLongs(1, 2, 3, 4, 5), AddressOf FindLongCallback, 0, 0&
            Test.Use NewSingles(1, 2, 3, 4, 5), AddressOf FindSingleCallback, 0, 0!
            Test.Use NewDoubles(1, 2, 3, 4, 5), AddressOf FindDoubleCallback, 0, 0#
            Test.Use NewCurrencys(1, 2, 3, 4, 5), AddressOf FindCurrencyCallback, 0, 0@
            Test.Use NewDates(#1/1/2001#), AddressOf FindDateCallback, #1/2/2001#, #12:00:00 AM#
            Test.Use ArrInt32, AddressOf FindInt32Callback, 3, Nothing
            
        Case "FindLastIndex_WithValues_ReturnsExpected"
            Test.Use NewLongs(1, 3, 3, 4, 5), 3, 2
            Test.Use NewLongs(1, 3, 3, 4, 5), 6, -1
            Test.Use Arr, 1, -2
        
        Case "FindAll_WithValues_ReturnsExpected"
            Test.Use NewLongs(1, 1, 2, 2, 2, 3, 3, 4), AddressOf FindLongCallback, 2, NewLongs(2, 2, 2)
            Test.Use NewLongs(1, 1, 2, 2, 2, 3, 3, 4), AddressOf FindLongCallback, 5, NewLongs()
            ReDim ArrInt32(3)
            Set ArrInt32(0) = MakeInt32(1)
            Set ArrInt32(1) = MakeInt32(2)
            Set ArrInt32(2) = MakeInt32(3)
            Set ArrInt32(3) = MakeInt32(2)
            Test.Use ArrInt32, AddressOf FindInt32Callback, 2, Array(ArrInt32(1), ArrInt32(3))
            
        Case "TrueForAll_WithValues_ReturnsExpected"
            Test.Use NewLongs(1, 1, 1), AddressOf FindLongCallback, 1, True
            Test.Use NewLongs(1, 2, 1), AddressOf FindLongCallback, 1, False
    End Select
End Sub

Public Sub FindIndex_WithValues_ReturnsExpected(ByRef Arr As Variant, ByVal ValueToFind As Long, ByVal Expected As Long)
    CorArrayCallbacks.FindCallbackValue = ValueToFind
    Dim Actual As Long
    
    Actual = CorArray.FindIndex(Arr, AddressOf FindLongCallback)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Find_WithValues_ReturnsExpected(ByRef Arr As Variant, ByVal AddressOfPredicate As Long, ByRef ValueToFind As Variant, ByRef Expected As Variant)
    Dim Actual As Variant
    CorArrayCallbacks.FindCallbackValue = ValueToFind
    
    MoveVariant Actual, CorArray.Find(Arr, AddressOfPredicate)
    
    Assert.That Actual, Iz.EqualTo(Expected).Strict
End Sub

Public Sub FindLastIndex_WithValues_ReturnsExpected(ByRef Arr As Variant, ByVal ValueToFind As Long, ByVal Expected As Long)
    CorArrayCallbacks.FindCallbackValue = ValueToFind
    Dim Actual As Long
    
    Actual = CorArray.FindLastIndex(Arr, AddressOf FindLongCallback)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FindAll_WithValues_ReturnsExpected(ByRef Arr As Variant, ByVal AddressOfPredicate As Long, ByRef ValueToFind, ByRef Expected As Variant)
    CorArrayCallbacks.FindCallbackValue = ValueToFind
    Dim Actual As Variant
    
    Actual = CorArray.FindAll(Arr, AddressOfPredicate)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TrueForAll_WithValues_ReturnsExpected(ByRef Arr As Variant, ByVal AddressOffPredicate As Long, ByRef ValueToTest, ByVal Expected As Boolean)
    CorArrayCallbacks.FindCallbackValue = ValueToTest
    Dim Actual As Variant
    
    Actual = CorArray.TrueForAll(Arr, AddressOffPredicate)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Find_WithUDT_ReturnsFoundUDT()
    Dim Arr(1) As VBGUID
    Arr(0).Data1 = 1
    Arr(1).Data1 = 2
    Dim Actual As VBGUID
    CorArrayCallbacks.FindCallbackValue = 2
    
    Actual = CorArray.Find(Arr, AddressOf FindVBGuidCallback)
    
    Assert.That Actual.Data1, Iz.EqualTo(2)
End Sub

Public Sub Find_WithUDT_ReturnsDefaultUDT()
    Dim Arr(1) As VBGUID
    Arr(0).Data1 = 1
    Arr(1).Data1 = 2
    Dim Actual As VBGUID
    CorArrayCallbacks.FindCallbackValue = 3
    
    Actual = CorArray.Find(Arr, AddressOf FindVBGuidCallback)
    
    Assert.That Actual.Data1, Iz.EqualTo(0)
End Sub

Public Sub FindIndex_WithIndex_FindsAfterIndex()
    Dim Arr() As Long
    Arr = NewLongs(2, 2, 3, 1, 2, 3)
    CorArrayCallbacks.FindCallbackValue = 2
    Dim Actual As Long
    
    Actual = CorArray.FindIndex(Arr, AddressOf FindLongCallback, 2)
    
    Const Expected As Long = 4
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FindIndex_WithIndexAndCount_DoesNotFindInRange()
    Dim Arr() As Long
    Arr = NewLongs(4, 2, 3, 1, 2, 4)
    CorArrayCallbacks.FindCallbackValue = 4
    Dim Actual As Long
    
    Actual = CorArray.FindIndex(Arr, AddressOf FindLongCallback, 1, 4)
    
    Const Expected As Long = -1
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FindIndex_WithNullArray_Throws()
    Dim Arr() As Long
    
    On Error Resume Next
    CorArray.FindIndex Arr, AddressOf FindLongCallback
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub FindIndex_WithIndexLessThanLBound_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.FindIndex Arr, AddressOf FindLongCallback, -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub FindIndex_WithCountLessThanZero_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.FindIndex Arr, AddressOf FindLongCallback, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub FindIndex_WithIndexAndCountPassedEndOfArray_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.FindIndex Arr, AddressOf FindLongCallback, 0, 2
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub Exists_WithExistingValue_ReturnsTrue()
    Dim Arr() As Long
    Arr = NewLongs(1, 2, 3, 4, 5)
    CorArrayCallbacks.FindCallbackValue = 4
    Dim Actual As Boolean
    
    Actual = CorArray.Exists(Arr, AddressOf FindLongCallback)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Exists_WithNonExistingValue_ReturnsFalse()
    Dim Arr() As Long
    Arr = NewLongs(1, 2, 3, 4, 5)
    CorArrayCallbacks.FindCallbackValue = 6
    Dim Actual As Boolean
    
    Actual = CorArray.Exists(Arr, AddressOf FindLongCallback)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub FindLastIndex_WithNullArray_Throws()
    Dim Arr() As Long
    
    On Error Resume Next
    CorArray.FindLastIndex Arr, AddressOf FindLongCallback
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub FindLastIndex_WithIndexLessThanLBounds_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.FindLastIndex Arr, AddressOf FindLongCallback, -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub FindLastIndex_WithCountLessThanZero_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.FindLastIndex Arr, AddressOf FindLongCallback, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub FindLastIndex_WithIndexAndCountPassedEndOfArray_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    CorArray.FindLastIndex Arr, AddressOf FindLongCallback, 0, 2
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub FindLast_WithExistingValue_ReturnsLastInstance()
    CorArrayCallbacks.FindCallbackValue = 2
    Dim Arr(2) As Int32
    Set Arr(0) = MakeInt32(2)
    Set Arr(1) = MakeInt32(2)
    Set Arr(2) = MakeInt32(3)
    Dim Actual As Int32
    
    Set Actual = CorArray.FindLast(Arr, AddressOf FindInt32Callback)
    
    Assert.That Actual, Iz.SameAs(Arr(1))
End Sub

Public Sub FindLast_WithNonExistingValue_ReturnsDefault()
    CorArrayCallbacks.FindCallbackValue = 4
    Dim Arr(2) As Int32
    Set Arr(0) = MakeInt32(2)
    Set Arr(1) = MakeInt32(2)
    Set Arr(2) = MakeInt32(3)
    Dim Actual As Int32
    
    Set Actual = CorArray.FindLast(Arr, AddressOf FindInt32Callback)
    
    Assert.That Actual, Iz.Nothing
End Sub

Public Sub FindAll_WithUDTs_ReturnsFoundUDTs()
    CorArrayCallbacks.FindCallbackValue = 2
    Dim Arr(2) As VBGUID
    Arr(1).Data1 = 2
    Arr(2).Data1 = 2
    Dim Actual() As VBGUID
    
    Actual = CorArray.FindAll(Arr, AddressOf FindVBGuidCallback)
    
    Assert.That UBound(Actual), Iz.EqualTo(1)
    Assert.That Actual(0).Data1, Iz.EqualTo(2)
    Assert.That Actual(1).Data1, Iz.EqualTo(2)
End Sub
