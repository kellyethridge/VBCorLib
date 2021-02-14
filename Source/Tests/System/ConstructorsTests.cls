VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstructorsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ConstructorsTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "NewChars_WithStringWithLengthNotEqualToOne_ThrowsFormatException"
            Test.Use ""
            Test.Use "aa"
        
        Case "NewStrings_WithValue_ReturnsStringArray"
            Test.Use "1", Array("1")
            Test.Use 2, Array("2")
            Test.Use CInt64(64), Array("64")
            
        Case "NewBigIntegers_WithValues_ReturnsBigIntegerArray"
            Test.Use 1, BigInteger.One
            Test.Use "1", BigInteger.One
            Test.Use Int64.MaxValue, NewBigInteger(Int64.MaxValue)
            
    End Select
End Sub

Public Sub NewDecimals_WithInt64_ReturnsDecimalArray()
    Dim Actual() As Variant
    
    Actual = NewDecimals(CInt64(1))

    Assert.That Actual, Iz.EqualTo(Array(CDec(1))).Strict
End Sub

Public Sub NewInt64s_WhenCalled_ReturnsInt64Array()
    Dim Actual() As Int64
    
    Actual = NewInt64s(1, 2, 3)
    
    Assert.That Actual, Iz.EqualTo(Array(CInt64(1), CInt64(2), CInt64(3)))
End Sub

Public Sub NewChars_WithOutOfRangeValue_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    NewChars &H10040
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub NewChars_WithStringWithLengthNotEqualToOne_ThrowsFormatException(ByVal Value As String)
    On Error Resume Next
    
    NewChars Value
    
    AssertFormatException Err
End Sub

Public Sub NewStrings_WithValue_ReturnsStringArray(ByRef Value As Variant, ByRef Expected As Variant)
    Dim Actual() As String
    
    Actual = NewStrings(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub NewBigIntegers_WithValues_ReturnsBigIntegerArray(ByRef Value As Variant, ByRef ExpectedElement As Variant)
    Dim Actual() As BigInteger
    Dim Expected() As BigInteger
    
    Actual = NewBigIntegers(Value)
    
    ReDim Expected(0)
    Set Expected(0) = CBigInt(ExpectedElement)
    Assert.That Actual, Iz.EqualTo(Expected).Using(New CorObjectComparer)
End Sub

