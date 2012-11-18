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
' TestcArrayFinds
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "FindIndex_WithValues_ReturnsExpected"
            Test.Use NewLongs(1, 2, 3, 4, 5), 3, 2
            Test.Use NewLongs(1, 2, 3, 4, 5), 6, -1
            
            Dim Arr(-1 To 1) As Long
            Test.Use Arr, 1, -2
        
        Case "Find_WithValues_ReturnsExpected"
            Dim ArrInt32(1) As Int32
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
    End Select
End Sub

Public Sub FindIndex_WithValues_ReturnsExpected(ByRef Arr As Variant, ByVal ValueToFind As Long, ByVal Expected As Long)
    cArrayCallbacks.FindCallbackValue = ValueToFind
    Dim Actual As Long
    
    Actual = cArray.FindIndex(Arr, AddressOf FindLongCallback)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Find_WithValues_ReturnsExpected(ByRef Arr As Variant, ByVal AddressOfPredicate As Long, ByRef ValueToFind As Variant, ByRef Expected As Variant)
    Dim Actual As Variant
    cArrayCallbacks.FindCallbackValue = ValueToFind
    
    MoveVariant Actual, cArray.find(Arr, AddressOfPredicate)
    
    Assert.That Actual, Iz.EqualTo(Expected).Strict
End Sub

Public Sub Find_WithUDT_ReturnsFoundUDT()
    Dim Arr(1) As VBGUID
    Arr(0).Data1 = 1
    Arr(1).Data1 = 2
    Dim Actual As VBGUID
    cArrayCallbacks.FindCallbackValue = 2
    
    Actual = cArray.find(Arr, AddressOf FindVBGuidCallback)
    
    Assert.That Actual.Data1, Iz.EqualTo(2)
End Sub

Public Sub Find_WithUDT_ReturnsDefaultUDT()
    Dim Arr(1) As VBGUID
    Arr(0).Data1 = 1
    Arr(1).Data1 = 2
    Dim Actual As VBGUID
    cArrayCallbacks.FindCallbackValue = 3
    
    Actual = cArray.find(Arr, AddressOf FindVBGuidCallback)
    
    Assert.That Actual.Data1, Iz.EqualTo(0)
End Sub

Public Sub FindIndex_WithIndex_FindsAfterIndex()
    Dim Arr() As Long
    Arr = NewLongs(2, 2, 3, 1, 2, 3)
    cArrayCallbacks.FindCallbackValue = 2
    Dim Actual As Long
    
    Actual = cArray.FindIndex(Arr, AddressOf FindLongCallback, 2)
    
    Const Expected As Long = 4
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FindIndex_WithIndexAndCount_DoesNotFindInRange()
    Dim Arr() As Long
    Arr = NewLongs(4, 2, 3, 1, 2, 4)
    cArrayCallbacks.FindCallbackValue = 4
    Dim Actual As Long
    
    Actual = cArray.FindIndex(Arr, AddressOf FindLongCallback, 1, 4)
    
    Const Expected As Long = -1
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FindIndex_WithNullArray_Throws()
    Dim Arr() As Long
    
    On Error Resume Next
    cArray.FindIndex Arr, AddressOf FindLongCallback
    
    AssertArgumentNullException Err, "Arr"
End Sub

Public Sub FindIndex_WithIndexLessThanLBound_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    cArray.FindIndex Arr, AddressOf FindLongCallback, -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub FindIndex_WithCountLessThanZero_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    cArray.FindIndex Arr, AddressOf FindLongCallback, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub FindIndex_WithIndexAndCountPassedEndOfArray_Throws()
    Dim Arr(0) As Long
    
    On Error Resume Next
    cArray.FindIndex Arr, AddressOf FindLongCallback, 0, 2
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub Exists_WithExistingValue_ReturnsTrue()
    Dim Arr() As Long
    Arr = NewLongs(1, 2, 3, 4, 5)
    cArrayCallbacks.FindCallbackValue = 4
    Dim Actual As Boolean
    
    Actual = cArray.Exists(Arr, AddressOf FindLongCallback)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Exists_WithNonExistingValue_ReturnsFalse()
    Dim Arr() As Long
    Arr = NewLongs(1, 2, 3, 4, 5)
    cArrayCallbacks.FindCallbackValue = 6
    Dim Actual As Boolean
    
    Actual = cArray.Exists(Arr, AddressOf FindLongCallback)
    
    Assert.That Actual, Iz.False
End Sub


'Public Sub TestFindAllObjects()
'    Dim a(1 To 50) As Int32
'    Dim i As Long
'
'    For i = 1 To 50
'        Set a(i) = New Int32
'        a(i).mValue = i
'    Next i
'    a(12).mValue = 32
'    a(15).mValue = 32
'    a(30).mValue = 32
'    a(45).mValue = 32
'
'    Dim o() As Int32
'    o = cArray.FindAll(a, AddressOf FindInt32)
'    Assert.AreEqual 5, cArray.GetLength(o)
'
'    Assert.AreEqual a(12), o(0)
'    Assert.AreEqual a(15), o(1)
'    Assert.AreEqual a(30), o(2)
'    Assert.AreEqual a(32), o(3)
'    Assert.AreEqual a(45), o(4)
'End Sub
'
'Public Sub TestFindNoObjects()
'    Dim a(10) As Int32
'    Dim i As Long
'
'    For i = 0 To 10
'        Set a(i) = New Int32
'    Next i
'
'    Dim o() As Int32
'    o = cArray.FindAll(a, AddressOf FindInt32)
'
'    Assert.AreEqual 0, cArray.GetLength(o)
'End Sub
'
'Public Sub TestFindLastIndex()
'    Dim a(10) As String
'    Dim i As Long
'
'    For i = 0 To 10
'        a(i) = "Item " & i
'    Next i
'
'    a(3) = "Hello"
'    a(4) = "Hello"
'    Assert.AreEqual 4, cArray.FindLastIndex(a, AddressOf FindHelloString)
'    Assert.AreEqual "Hello", cArray.FindLast(a, AddressOf FindHelloString)
'
'    a(3) = "Bye"
'    a(4) = "Bye"
'    Assert.AreEqual -1, cArray.FindLastIndex(a, AddressOf FindHelloString)
'    Assert.AreEqual "", cArray.FindLast(a, AddressOf FindHelloString)
'End Sub
'
'Public Sub TestTrueForAll()
'    Dim a(10) As String
'    Dim i As Long
'
'    For i = 0 To 10
'        a(i) = "Hello"
'    Next i
'
'    Assert.IsTrue cArray.TrueForAll(a, AddressOf FindHelloString)
'
'    a(4) = "Bye"
'    Assert.IsFalse cArray.TrueForAll(a, AddressOf FindHelloString)
'End Sub

'Public Sub TestFindAllUDTs()
'    Dim a(20) As VBGUID
'    Dim i As Long
'    Dim g As Variant
'
'    a(5).Data1 = 7
'    a(9).Data1 = 7
'    a(10).Data1 = 7
'    a(16).Data1 = 7
'    a(20).Data1 = 7
'
'    g = cArray.FindAll(a, AddressOf FindVBGuid)
'    Assert.AreEqual 5, cArray.GetLength(g)
'
'    For i = 0 To 4
'        Assert.AreEqual 7, g(i).Data1
'    Next i
'End Sub
