VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   ObjectTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetHashCode_WithMinimumValues_ReturnsZero"
            Test.Use Nothing
            Test.Use Null
            Test.Use Empty
            Test.Use Missing
            Test.Use 0#
            Test.Use 0
            Test.Use cDateTime.MinValue
            Test.Use TimeSpan.Zero
            Test.Use CDec(0)
            
        Case "GetHashCode_WithSameValueTwice_ReturnsSameValue"
            Test.Use 1
            Test.Use 1#
            Test.Use #1/1/2001#
            Test.Use 1234.1234@
            Test.Use CDec("123421346876234143.2343243")
            Test.Use "abc"
        
        Case "Equals_WithValues_ReturnsExpected"
            Test.Use Nothing, Nothing, True
            Test.Use Nothing, "a", False
            Test.Use "a", Nothing, False
            Test.Use 1, 1, True
            Test.Use 1&, 1, False
            Test.Use 1, 1&, False
            Test.Use "a", "a", True
            Test.Use "A", "a", False
            Test.Use 1.1, 1.1, True
            Test.Use 1.1@, 1.1@, True
            Test.Use 1.1@, 1.1, False
            Test.Use 1.1!, 1.1!, True
            Test.Use 1.1!, 1.1, False
            Test.Use #1/1/2001#, #1/1/2001#, True
            Test.Use NewDate(#1/1/2001#), NewDate(#1/1/2001#), True
            Test.Use NewDate(#1/1/2001#), #1/1/2001#, True
            Test.Use #1/1/2001#, NewDate(#1/1/2001#), True
            
    End Select
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByRef ObjA As Variant, ByRef ObjB As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Object.Equals(ObjA, ObjB)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetHashCode_WithMinimumValues_ReturnsZero(ByRef Value As Variant)
    Dim Actual As Long
    
    Actual = Object.GetHashCode(Value)
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

Public Sub GetHashCode_WithSameValueTwice_ReturnsSameValue(ByRef Value As Variant)
    Dim Actual As Boolean
    
    Actual = Object.GetHashCode(Value) = Object.GetHashCode(Value)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub GetHashCode_WithTwoReferencesToSameObjectInstance_ReturnsSameValue()
    Dim Instance1 As New Collection
    Dim Instance2 As Collection
    Set Instance2 = Instance1
    
    Dim Actual As Boolean
    Actual = Object.GetHashCode(Instance1) = Object.GetHashCode(Instance2)
    
    Assert.That Actual, Iz.True
End Sub
