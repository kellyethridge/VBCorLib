VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MathExtTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Ceiling_WithValues_ReturnsExpected"
            Test.Use 2.9, 3#
            Test.Use -2.9, -2#
            Test.Use 3#, 3#
            Test.Use 3.1, 4#
            Test.Use -3.1, -3#
        
        Case "Floor_WithValues_ReturnsExpected"
            Test.Use 2.1, 2#
            Test.Use 2.9, 2#
            Test.Use 0.9, 0#
            Test.Use -0.9, -1#
            Test.Use -0.1, -1#
            Test.Use -1.1, -2#
            
        Case "Max_WithValues_ReturnsExpected"
            Test.Use 5, 5, 5
            Test.Use 5, 2, 5
            Test.Use 2, 5, 5
            Test.Use NewInt32(2), NewInt32(5), NewInt32(5)
            Test.Use NewInt32(5), NewInt32(2), NewInt32(5)
        
        Case "Min_WithValues_ReturnsExpected"
            Test.Use 5, 5, 5
            Test.Use 5, 2, 2
            Test.Use 2, 5, 2
            Test.Use NewInt32(2), NewInt32(5), NewInt32(2)
            Test.Use NewInt32(5), NewInt32(2), NewInt32(2)
            
        Case "DivRem_WithValues_ReturnsExpected"
            Test.Use 11, 4, 2, 3
            Test.Use -11, 4, -2, -3
            Test.Use -11, -4, 2, -3
            Test.Use 11, -4, -2, 3
        
        Case "MinIn_WithValues_ReturnsExpected"
            Test.Use Array(5, 4, 3, 2, 1), 1
            Test.Use Array(#1/1/2001#, #1/2/2001#, #1/3/2001#), #1/1/2001#
            Test.Use Array(NewDate(#1/1/2001#), NewDate(#1/2/2001#), NewDate(#1/1/2000#)), NewDate(#1/1/2000#)
            Test.Use Array(), Empty
            Test.Use NewCollection(5, 4, 3, 2, 1), 1
            Test.Use NewCollection(#1/1/2001#, #1/2/2001#, #1/3/2001#), #1/1/2001#
            Test.Use NewCollection(NewDate(#1/1/2001#), NewDate(#1/2/2001#), NewDate(#1/1/2000#)), NewDate(#1/1/2000#)
            Test.Use New Collection, Empty
            Test.Use Cor.NewArrayList(Source:=Array(5, 4, 3, 2, 1)), 1
            Test.Use Cor.NewArrayList(Source:=Array(#1/1/2001#, #1/2/2001#, #1/3/2001#)), #1/1/2001#
            Test.Use Cor.NewArrayList(Source:=Array(NewDate(#1/1/2001#), NewDate(#1/2/2001#), NewDate(#1/1/2000#))), NewDate(#1/1/2000#)
            Test.Use New ArrayList, Empty
            
        Case "MaxIn_WithValues_ReturnsExpected"
            Test.Use Array(5, 4, 3, 2, 1), 5
            Test.Use Array(#1/1/2001#, #1/2/2001#, #1/3/2001#), #1/3/2001#
            Test.Use Array(NewDate(#1/1/2001#), NewDate(#1/2/2001#), NewDate(#1/1/2000#)), NewDate(#1/2/2001#)
            Test.Use Array(), Empty
            Test.Use NewCollection(5, 4, 3, 2, 1), 5
            Test.Use NewCollection(#1/1/2001#, #1/2/2001#, #1/3/2001#), #1/3/2001#
            Test.Use NewCollection(NewDate(#1/1/2001#), NewDate(#1/2/2001#), NewDate(#1/1/2000#)), NewDate(#1/2/2001#)
            Test.Use New Collection, Empty
            Test.Use Cor.NewArrayList(Source:=Array(5, 4, 3, 2, 1)), 5
            Test.Use Cor.NewArrayList(Source:=Array(#1/1/2001#, #1/2/2001#, #1/3/2001#)), #1/3/2001#
            Test.Use Cor.NewArrayList(Source:=Array(NewDate(#1/1/2001#), NewDate(#1/2/2001#), NewDate(#1/1/2000#))), NewDate(#1/2/2001#)
            Test.Use New ArrayList, Empty
        
    End Select
End Sub

Public Sub Ceiling_WithValues_ReturnsExpected(ByVal Value As Double, ByVal Expected As Double)
    Dim Actual As Double
    
    Actual = Ceiling(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Floor_WithValues_ReturnsExpected(ByVal Value As Double, ByVal Expected As Double)
    Dim Actual As Double
    
    Actual = Floor(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ACos_WhenCalled_ReturnsExpected()
    Dim Actual As Double
    
    Actual = ACos(0.23456)
    
    Assert.That Actual, Iz.EqualTo(1.33403041100045).Within(0.000000001)
End Sub

Public Sub Max_WithValues_ReturnsExpected(ByRef Value1 As Variant, ByRef Value2 As Variant, ByRef Expected As Variant)
    Dim Actual As Variant
    
    MoveVariant Actual, Max(Value1, Value2)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Min_WithValues_ReturnsExpected(ByRef Value1 As Variant, ByRef Value2 As Variant, ByRef Expected As Variant)
    Dim Actual As Variant
    
    MoveVariant Actual, Min(Value1, Value2)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Max_WithNonComparable_ThrowsArgumentException()
    On Error Resume Next
    
    Max Me, New MathExtTests
    
    AssertArgumentException Err
End Sub

Public Sub DivRem_WithValues_ReturnsExpected(ByVal a As Long, ByVal b As Long, ByVal ExpectedQ As Long, ByVal ExpectedR As Long)
    Dim ActualQ As Long
    Dim ActualR As Long
    
    ActualQ = DivRem(a, b, ActualR)
    
    Assert.That ActualQ, Iz.EqualTo(ExpectedQ)
    Assert.That ActualR, Iz.EqualTo(ExpectedR)
End Sub

Public Sub MinIn_WithValues_ReturnsExpected(ByRef List As Variant, ByRef Expected As Variant)
    Dim Actual As Variant
    
    MoveVariant Actual, MinIn(List)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub MaxIn_WithValues_ReturnsExpected(ByRef List As Variant, ByRef Expected As Variant)
    Dim Actual As Variant
    
    MoveVariant Actual, MaxIn(List)
    
    Assert.That Actual, Equals(Expected)
End Sub


