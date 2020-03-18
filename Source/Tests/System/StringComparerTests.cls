VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringComparerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' StringComparerTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Compare_WithOrdinalComparer_ReturnsExpected"
            Test.Use("a", "a").Expect Iz.EqualTo(0)
            Test.Use("a", "b").Expect Iz.LessThan(0)
            Test.Use("b", "a").Expect Iz.GreaterThan(0)
            Test.Use("a", "A").Expect Iz.GreaterThan(0)
            Test.Use("A", "a").Expect Iz.LessThan(0)
            Test.Use("a", "B").Expect Iz.GreaterThan(0)
            Test.Use("A", "b").Expect Iz.LessThan(0)
            Test.Use(1, 1).Expect Iz.EqualTo(0)
            Test.Use(1, 0).Expect Iz.GreaterThan(0)
            Test.Use(0, 1).Expect Iz.LessThan(0)
        
        Case "Compare_WithOrdinalIgnoreCaseComparer_ReturnsExpected"
            Test.Use("a", "a").Expect Iz.EqualTo(0)
            Test.Use("a", "b").Expect Iz.LessThan(0)
            Test.Use("b", "a").Expect Iz.GreaterThan(0)
            Test.Use("a", "A").Expect Iz.EqualTo(0)
            Test.Use("A", "a").Expect Iz.EqualTo(0)
            Test.Use("a", "B").Expect Iz.LessThan(0)
            Test.Use("A", "b").Expect Iz.LessThan(0)
            Test.Use(1, 1).Expect Iz.EqualTo(0)
            Test.Use(1, 0).Expect Iz.GreaterThan(0)
            Test.Use(0, 1).Expect Iz.LessThan(0)
            
        Case "Equals_WithOrdinalComparer_ReturnsExpected"
            Test.Use "a", "a", True
            Test.Use "a", "b", False
            Test.Use "a", "A", False
            Test.Use "A", "a", False
            Test.Use 1, 1, True
            Test.Use 1, 0, False
            
        Case "Equals_WithOrdinalIgnoreCaseComparer_ReturnsExpected"
            Test.Use "a", "a", True
            Test.Use "a", "b", False
            Test.Use "a", "A", True
            Test.Use "A", "a", True
            Test.Use 1, 1, True
            Test.Use 1, 0, False
            
        Case "GetHashCode_WithOrdinalComparer_ReturnsExpected"
            Test.Use "a", "a", True
            Test.Use "a", "A", False
            Test.Use 1, 1, True
            Test.Use 1, 0, False
            
        Case "GetHashCode_WithOrdinalIgnoreCaseComparer_ReturnsExpected"
            Test.Use "a", "a", True
            Test.Use "a", "A", True
            Test.Use "A", "a", True
            Test.Use "a", "b", False
            Test.Use 1, 1, True
            Test.Use 1, 0, False
            
    End Select
End Sub

Public Function Compare_WithOrdinalComparer_ReturnsExpected(ByRef x As Variant, ByRef y As Variant) As Long
    Dim Actual As Long
    
    Actual = StringComparer.Ordinal.Compare(x, y)
    
    Compare_WithOrdinalComparer_ReturnsExpected = Actual
End Function

Public Function Compare_WithOrdinalIgnoreCaseComparer_ReturnsExpected(ByRef x As Variant, ByRef y As Variant) As Long
    Dim Actual As Long
    
    Actual = StringComparer.OrdinalIgnoreCase.Compare(x, y)
    
    Compare_WithOrdinalIgnoreCaseComparer_ReturnsExpected = Actual
End Function

Public Sub Equals_WithOrdinalComparer_ReturnsExpected(ByRef x As Variant, ByRef y As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = StringComparer.Ordinal.Equals(x, y)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithOrdinalIgnoreCaseComparer_ReturnsExpected(ByRef x As Variant, ByRef y As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = StringComparer.OrdinalIgnoreCase.Equals(x, y)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetHashCode_WithOrdinalComparer_ReturnsExpected(ByRef x As Variant, ByRef y As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = (StringComparer.Ordinal.GetHashCode(x) = StringComparer.Ordinal.GetHashCode(y))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetHashCode_WithOrdinalIgnoreCaseComparer_ReturnsExpected(ByRef x As Variant, ByRef y As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = (StringComparer.OrdinalIgnoreCase.GetHashCode(x) = StringComparer.OrdinalIgnoreCase.GetHashCode(y))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithTextCompare_ReturnsExpected()
    Dim Actual As Boolean
    
    Actual = StringComparer.TextCompare.Equals("a", "A")
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Compare_WithTextCompare_ReturnsExpected()
    Assert.That StringComparer.TextCompare.Compare("a", "A"), Iz.EqualTo(0)
    Assert.That StringComparer.TextCompare.Compare("a", "b"), Iz.EqualTo(-1)
    Assert.That StringComparer.TextCompare.Compare("b", "a"), Iz.EqualTo(1)
End Sub

Public Sub Compare_WithBinaryCompare_ReturnsExpected()
    Assert.That StringComparer.BinaryCompare.Compare("a", "a"), Iz.EqualTo(0)
    Assert.That StringComparer.BinaryCompare.Compare("a", "A"), Iz.EqualTo(1)
End Sub

Public Sub Equals_WithBinaryCompare_ReturnsExpected()
    Assert.That StringComparer.BinaryCompare.Equals("a", "a"), Iz.True
    Assert.That StringComparer.BinaryCompare.Equals("a", "A"), Iz.False
End Sub

Public Sub Equals_WithInvariantCulture_ReturnsExpected()
    Assert.That StringComparer.InvariantCulture.Equals("a", "a"), Iz.True
    Assert.That StringComparer.InvariantCulture.Equals("a", "A"), Iz.False
End Sub

Public Sub Compare_WithInvariantCulture_ReturnsExpected()
    Assert.That StringComparer.InvariantCulture.Compare("a", "a"), Iz.EqualTo(0)
    Assert.That StringComparer.InvariantCulture.Compare("a", "A"), Iz.EqualTo(-1)
End Sub

Public Sub Equals_WithInvariantCultureIgnoreCase_ReturnsExpected()
    Assert.That StringComparer.InvariantCultureIgnoreCase.Equals("a", "A"), Iz.True
End Sub

Public Sub Equals_WithCurrentCulture_ReturnsExpected()
    Assert.That StringComparer.CurrentCulture.Equals("a", "a"), Iz.True
End Sub

Public Sub Equals_WithCurrentCultureIgnoreCase_ReturnsExpected()
    Assert.That StringComparer.CurrentCultureIgnoreCase.Equals("a", "A"), Iz.True
End Sub

