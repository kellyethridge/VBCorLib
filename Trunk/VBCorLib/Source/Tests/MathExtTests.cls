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

Public Sub TestMin()
    Assert.AreEqual 2, Min(2, 5)
    Assert.AreEqual 2, Min(5, 2)
End Sub

Public Sub TestMinObject()
    Dim X As New Int32
    Dim Y As New Int32
    
    X.Value = 2
    Y.Value = 5
    
    Assert.AreEqual X, Min(X, Y)
    Assert.AreEqual X, Min(Y, X)
End Sub

Public Sub TestDivRem()
    Dim Q As Long
    Dim r As Long
    
    Q = DivRem(11, 4, r)
    
    Assert.AreEqual 2, Q
    Assert.AreEqual 3, r
End Sub

Public Sub TestDivRemNegNum()
    Dim Q As Long
    Dim r As Long
    
    Q = DivRem(-11, 4, r)
    
    Assert.AreEqual -2, Q
    Assert.AreEqual -3, r
End Sub

Public Sub TestDivRemNegNumAndDenom()
    Dim Q As Long
    Dim r As Long
    
    Q = DivRem(-11, -4, r)
    
    Assert.AreEqual 2, Q
    Assert.AreEqual -3, r
End Sub

Public Sub TestDivRemNegDenom()
    Dim Q As Long
    Dim r As Long
    
    Q = DivRem(11, -4, r)
    
    Assert.AreEqual -2, Q
    Assert.AreEqual 3, r
End Sub

Public Sub TestMinInArray()
    Dim a(10)
    Dim i As Long
    For i = 0 To 10
        a(i) = 10 - i
    Next i
    
    Assert.AreEqual 0, MinIn(a)
End Sub

Public Sub TestMaxInArray()
    Dim a(10)
    Dim i As Long
    For i = 0 To 10
        a(i) = i
    Next i
    
    Assert.AreEqual 10, MaxIn(a)
End Sub

