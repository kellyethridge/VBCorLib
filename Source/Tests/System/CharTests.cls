VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CharTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CharTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "IsWhiteSpace_WithString_ReturnsExpected"
            Test.Use ChrW$(&H20), True
            Test.Use ChrW$(&H1680), True
            Test.Use ChrW$(&H180E), True
            Test.Use ChrW$(&H2000), True
            Test.Use ChrW$(&H2001), True
            Test.Use ChrW$(&H2002), True
            Test.Use ChrW$(&H2003), True
            Test.Use ChrW$(&H2004), True
            Test.Use ChrW$(&H2005), True
            Test.Use ChrW$(&H2006), True
            Test.Use ChrW$(&H2007), True
            Test.Use ChrW$(&H2008), True
            Test.Use ChrW$(&H2009), True
            Test.Use ChrW$(&H200A), True
            Test.Use ChrW$(&H202F), True
            Test.Use ChrW$(&H205F), True
            Test.Use ChrW$(&H3000), True
            Test.Use ChrW$(&H2028), True
            Test.Use ChrW$(&H2029), True
            Test.Use ChrW$(&H9), True
            Test.Use ChrW$(&HA), True
            Test.Use ChrW$(&HB), True
            Test.Use ChrW$(&HC), True
            Test.Use ChrW$(&HD), True
            Test.Use ChrW$(&H85), True
            Test.Use ChrW$(&HA0), True
            Test.Use ChrW$(65), False
            
        Case "IsWhiteSpace_WithStringIndex_ReturnsExpected"
            Test.Use 0, True
            Test.Use 1, False
            Test.Use 2, True
            Test.Use 3, False
            Test.Use 4, True
            
        Case "IsWhiteSpaceChar_WithChar_ReturnsExpected"
            Test.Use &H20, True
            Test.Use &H1680, True
            Test.Use &H180E, True
            Test.Use &H2000, True
            Test.Use &H2001, True
            Test.Use &H2002, True
            Test.Use &H2003, True
            Test.Use &H2004, True
            Test.Use &H2005, True
            Test.Use &H2006, True
            Test.Use &H2007, True
            Test.Use &H2008, True
            Test.Use &H2009, True
            Test.Use &H200A, True
            Test.Use &H202F, True
            Test.Use &H205F, True
            Test.Use &H3000, True
            Test.Use &H2028, True
            Test.Use &H2029, True
            Test.Use &H9, True
            Test.Use &HA, True
            Test.Use &HB, True
            Test.Use &HC, True
            Test.Use &HD, True
            Test.Use &H85, True
            Test.Use &HA0, True
            Test.Use 65, False
            
        Case "IsHighSurrogateChar"
            Test.Use &HD800&, True
            Test.Use &HD800, True
            Test.Use &HD7FF&, False
            Test.Use &HD7FF, False
            Test.Use &HDBFF&, True
            Test.Use &HDBFF, True
            Test.Use &HDC00&, False
            Test.Use &HDC00, False
            
        Case "IsLowSurrogateChar"
            Test.Use &HDC00, True
            Test.Use &HDC00&, True
            Test.Use &HDFFF, True
            Test.Use &HDFFF&, True
            Test.Use &HDBFF, False
            Test.Use &HDBFF&, False
            Test.Use &HE000, False
            Test.Use &HE000&, False
            
        Case "IsHighSurrogate_WithValues_ReturnsExpected"
            Test.Use "a", 0, False
            Test.Use ChrW$(&HD800), 0, True
            Test.Use "aa", 1, False
            Test.Use "a" & ChrW$(&HDBFF), 1, True
            
        Case "IsLowSurrogate_WithValues_ReturnsExpected"
            Test.Use ChrW$(&HD800), 0, False
            Test.Use ChrW$(&HDC00), 0, True
            Test.Use "a" & ChrW$(&HDBFF), 1, False
            Test.Use "a" & ChrW$(&HDFFF), 1, True
            
    End Select
End Sub

Public Sub IsWhiteSpace_WithString_ReturnsExpected(ByVal s As String, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsWhiteSpace(s, 0)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsWhiteSpace_WithStringIndex_ReturnsExpected(ByVal Index As Long, ByVal Expected As Boolean)
    Const s As String = " a b "
    Dim Actual As Boolean
     
    Actual = Char.IsWhiteSpace(s, Index)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsWhiteSpaceChar_WithChar_ReturnsExpected(ByVal Ch As Integer, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsWhiteSpaceChar(Ch)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsWhiteSpace_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsWhiteSpace "a", -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsWhiteSpace_WithIndexPassedStringLength_ThrowsArgumentOutOfException()
    On Error Resume Next
    
    Char.IsWhiteSpace "a", 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsWhiteSpace_WithEmptyString_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsWhiteSpace "", 0
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub IsHighSurrogateChar(ByVal Value As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsHighSurrogateChar(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLowSurrogateChar(ByVal Value As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsLowSurrogateChar(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsHighSurrogate_WithValues_ReturnsExpected(ByRef s As String, ByVal Index As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsHighSurrogate(s, Index)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsHighSurrogate_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsHighSurrogate "a", -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsHighSurrogate_WithIndexPassedStringLength_ThrowsArgumentOutOfException()
    On Error Resume Next
    
    Char.IsHighSurrogate "a", 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsLowSurrogate_WithValues_ReturnsExpected(ByRef s As String, ByVal Index As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsLowSurrogate(s, Index)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLowSurrogate_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsLowSurrogate "a", -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsLowSurrogate_WithIndexPassedStringLength_ThrowsArgumentOutOfException()
    On Error Resume Next
    
    Char.IsLowSurrogate "a", 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub
