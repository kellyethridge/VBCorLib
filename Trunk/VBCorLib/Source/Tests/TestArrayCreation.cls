VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArrayCreateInstanceTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "NewXxx_WithValues_ReturnsExpected"
            Test.Use NewBytes(1, 2, 3), vbByte, Array(CByte(1), CByte(2), CByte(3))
            Test.Use NewIntegers(1, 2, 3), vbInteger, Array(1, 2, 3)
            Test.Use NewLongs(1, 2, 3), vbLong, Array(1&, 2&, 3&)
            Test.Use NewSingles(1, 2, 3), vbSingle, Array(1!, 2!, 3!)
            Test.Use NewDoubles(1, 2, 3), vbDouble, Array(1#, 2#, 3#)
            Test.Use NewCurrencys(1, 2, 3), vbCurrency, Array(1@, 2@, 3@)
            Test.Use NewBooleans(True, False), vbBoolean, Array(True, False)
            Test.Use NewDates(#1/1/2001#, #1/2/2001#), vbDate, Array(#1/1/2001#, #1/2/2001#)
            Test.Use NewDecimals(1, 2), vbVariant, Array(CDec(1), CDec(2))
            Test.Use NewStrings("1", "2", "3"), vbString, Array("1", "2", "3")
            
        Case "CreateInstance_WithOneDimension_CreatesExpected"
            Test.Use VbVarType.vbByte, 0, NewBytes()
            Test.Use VbVarType.vbByte, 1, NewBytes(0)
            Test.Use VbVarType.vbInteger, 1, NewIntegers(0)
            Test.Use VbVarType.vbLong, 1, NewLongs(0)
            Test.Use VbVarType.vbLong, 0, NewLongs()
            Test.Use VbVarType.vbSingle, 1, NewSingles(0)
            Test.Use VbVarType.vbDouble, 1, NewDoubles(0)
            Test.Use VbVarType.vbCurrency, 1, NewCurrencys(0)
            Test.Use VbVarType.vbDate, 1, NewDates(#12:00:00 AM#)
            Test.Use VbVarType.vbString, 1, NewStrings("")
            Test.Use VbVarType.vbBoolean, 1, NewBooleans(False)
            Test.Use VbVarType.vbVariant, 1, NewVariants(Empty)
            
        Case "NewArray_WithValues_ReturnsExpected"
            Test.Use VbVarType.vbByte, 1, NewBytes(1)
            Test.Use VbVarType.vbInteger, 1, NewIntegers(1)
            Test.Use VbVarType.vbLong, 1, NewLongs(1)
            Test.Use VbVarType.vbSingle, 1, NewSingles(1)
            Test.Use VbVarType.vbDouble, 1, NewDoubles(1)
            Test.Use VbVarType.vbString, "1", NewStrings("1")
            
        Case "CreateInstance_WithInvalidVbVarType_Throws"
            Test.Use VbVarType.vbArray
            Test.Use VbVarType.vbDataObject
            Test.Use VbVarType.vbEmpty
            Test.Use VbVarType.vbError
            Test.Use VbVarType.vbNull
            Test.Use VbVarType.vbUserDefinedType
            
    End Select
End Sub

Public Sub NewXxx_WithValues_ReturnsExpected(ByRef Actual As Variant, ByVal ExpectedType As VbVarType, ByRef Expected As Variant)
    Assert.That VarType(Actual), Iz.EqualTo(ExpectedType Or vbArray)
    Assert.That Actual, Iz.EqualTo(Expected).Strict
End Sub

Public Sub CreateInstance_WithOneDimension_CreatesExpected(ByVal vbType As VbVarType, ByVal Length As Long, ByRef Expected As Variant)
    Dim Actual As Variant
    
    Actual = cArray.CreateInstance(vbType, Length)
    
    Assert.That Actual, Iz.EqualTo(Expected).Strict
End Sub

Public Sub NewArray_WithValues_ReturnsExpected(ByVal ArrayType As VbVarType, ByRef Data As Variant, ByRef Expected As Variant)
    Dim Actual As Variant
    
    Actual = NewArray(ArrayType, Data)
    
    Assert.That Actual, Iz.EqualTo(Expected).Strict
End Sub

Public Sub CreateInstance_WithInvalidVbVarType_Throws(ByVal ElementType As VbVarType)
    On Error Resume Next
    
    cArray.CreateInstance ElementType
    
    AssertArgumentException Err, "ElementType"
End Sub

Public Sub CreateInstance_WithTwoDimensions_CreatesExpected()
    Dim Actual() As Long
    
    Actual = cArray.CreateInstance(vbLong, 1, 2)
    
    Dim Expected() As Long
    ReDim Expected(0, 1)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CreateInstance_WithThreeDimensions_CreatesExpected()
    Dim Actual() As Long
    
    Actual = cArray.CreateInstance(vbLong, 1, 2, 3)
    
    Dim Expected() As Long
    ReDim Expected(0, 1, 2)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CreateInstance_WithLength1LessThanZero_Throws()
    On Error Resume Next
    
    cArray.CreateInstance vbLong, -1
    
    AssertArgumentOutOfRangeException Err, "Length1"
End Sub

Public Sub CreateInstance_WithZeroOneAndTwoLength_CreatesExpected()
    Dim Actual() As Long
    
    Actual = cArray.CreateInstance(vbLong, 0, 1, 2)
    
    Assert.That UBound(Actual, 1), Iz.EqualTo(-1)
    Assert.That UBound(Actual, 2), Iz.EqualTo(0)
    Assert.That UBound(Actual, 3), Iz.EqualTo(1)
End Sub
