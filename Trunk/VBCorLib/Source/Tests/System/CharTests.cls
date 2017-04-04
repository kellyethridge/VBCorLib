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
        Case "IsWhiteSpaceStr_WithString_ReturnsExpected"
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
            
        Case "IsWhiteSpaceStr_WithStringIndex_ReturnsExpected"
            Test.Use 0, True
            Test.Use 1, False
            Test.Use 2, True
            Test.Use 3, False
            Test.Use 4, True
            
        Case "IsWhiteSpace_WithChar_ReturnsExpected"
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
            
        Case "IsSurrogate_WithValues_ReturnsExpected"
            Test.Use 0, False
            Test.Use &HD7FF, False
            Test.Use &HD7FF&, False
            Test.Use &HD800, True
            Test.Use &HD800&, True
            Test.Use &HDFFF, True
            Test.Use &HDFFF&, True
            Test.Use &HE000, False
            Test.Use &HE000&, False
            Test.Use &HFFFF, False
            
        Case "IsSurrogateStr_WithValues_ReturnsExpected"
            Test.Use "a", False
            Test.Use ChrW$(&HD7FF), False
            Test.Use ChrW$(&HD800), True
            Test.Use ChrW$(&HDFFF), True
            Test.Use ChrW$(&HE000), False
            Test.Use ChrW$(&HFFFF), False
            
        Case "IsHighSurrogate_WithValues_ReturnsExpected"
            Test.Use &HD800&, True
            Test.Use &HD800, True
            Test.Use &HD7FF&, False
            Test.Use &HD7FF, False
            Test.Use &HDBFF&, True
            Test.Use &HDBFF, True
            Test.Use &HDC00&, False
            Test.Use &HDC00, False
            
        Case "IsLowSurrogate_WithValues_ReturnsExpected"
            Test.Use &HDC00, True
            Test.Use &HDC00&, True
            Test.Use &HDFFF, True
            Test.Use &HDFFF&, True
            Test.Use &HDBFF, False
            Test.Use &HDBFF&, False
            Test.Use &HE000, False
            Test.Use &HE000&, False
            
        Case "IsHighSurrogateStr_WithValues_ReturnsExpected"
            Test.Use "a", 0, False
            Test.Use ChrW$(&HD800), 0, True
            Test.Use "aa", 1, False
            Test.Use "a" & ChrW$(&HDBFF), 1, True
            
        Case "IsLowSurrogateStr_WithValues_ReturnsExpected"
            Test.Use ChrW$(&HD800), 0, False
            Test.Use ChrW$(&HDC00), 0, True
            Test.Use "a" & ChrW$(&HDBFF), 1, False
            Test.Use "a" & ChrW$(&HDFFF), 1, True
            
        Case "Equals_WithValues_ReturnsExpected"
            Test.Use 1, 2, False
            Test.Use 1, 1, True
            Test.Use &HFFFF, &HFFFF&, True
            Test.Use &HD800&, &HD800, True
            Test.Use &HFFFFD800, &HD800, True
            Test.Use &HFFFFD800, &HD800&, True
            
        Case "Compare_WithValues_ReturnsExpected"
            Test.Use(1, 2).Expect Iz.Negative
            Test.Use(2, 1).Expect Iz.Positive
            Test.Use(2, 2).Expect Iz.EqualTo(0)
            Test.Use(&HD800, &HD800&).Expect Iz.EqualTo(0)
            Test.Use(&HD801, &HD800&).Expect Iz.Positive
            Test.Use(&HD801&, &HD800).Expect Iz.Positive
            Test.Use(&HFFFFD800, &HD800).Expect Iz.EqualTo(0)
            Test.Use(&HFFFFD800, &HD800&).Expect Iz.EqualTo(0)
            Test.Use(&HFFFFD801, &HD800).Expect Iz.Positive
            Test.Use(&HFFFFD800, &HD801).Expect Iz.Negative
            Test.Use(&HFFFFD800, &HD801&).Expect Iz.Negative
            Test.Use(&HFFFFD801, &HD800&).Expect Iz.Positive
            
        Case "ConvertToUtf32_WithValues_ReturnsExpected"
            Test.Use &HD800, &HDC00, &H10000
            Test.Use &HD800&, &HDC00&, &H10000
            Test.Use &HDBFF, &HDFFF, &H10FFFF
            Test.Use &HDBFF&, &HDFFF&, &H10FFFF
        
        Case "ConvertToUtf32Str_WithValues_ReturnsExpected"
            Test.Use t("\x0"), 0
            Test.Use t("\xD7FF"), &HD7FF&
            Test.Use t("\xE000"), &HE000&
            Test.Use t("\xFFFF"), &HFFFF&
            Test.Use t("\xD800\xDC00"), &H10000
            Test.Use t("\xDBFF\xDFFF"), &H10FFFF
            
    End Select
End Sub

Public Sub IsWhiteSpaceStr_WithString_ReturnsExpected(ByVal s As String, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsWhiteSpaceStr(s, 0)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsWhiteSpaceStr_WithStringIndex_ReturnsExpected(ByVal Index As Long, ByVal Expected As Boolean)
    Const s As String = " a b "
    Dim Actual As Boolean
     
    Actual = Char.IsWhiteSpaceStr(s, Index)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsWhiteSpace_WithChar_ReturnsExpected(ByVal Ch As Integer, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsWhiteSpace(Ch)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsWhiteSpaceStr_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsWhiteSpaceStr "a", -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsWhiteSpaceStr_WithIndexPassedStringLength_ThrowsArgumentOutOfException()
    On Error Resume Next
    
    Char.IsWhiteSpaceStr "a", 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsWhiteSpaceStr_WithEmptyString_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsWhiteSpaceStr "", 0
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub IsSurrogate_WithValues_ReturnsExpected(ByVal Value As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsSurrogate(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsSurrogateStr_WithValues_ReturnsExpected(ByRef Value As String, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsSurrogateStr(Value, 0)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsSurrogateStr_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsSurrogateStr "a", -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsSurrogateStr_WithIndexEqualToLengthOfString_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsSurrogateStr "abc", 3
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsSurrogateString_WithIndexGreaterThanLengthOfString_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsSurrogateStr "abc", 4
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsHighSurrogate_WithValues_ReturnsExpected(ByVal Value As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsHighSurrogate(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLowSurrogate_WithValues_ReturnsExpected(ByVal Value As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsLowSurrogate(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsHighSurrogateStr_WithValues_ReturnsExpected(ByRef s As String, ByVal Index As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsHighSurrogateStr(s, Index)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsHighSurrogateStr_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsHighSurrogateStr "a", -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsHighSurrogateStr_WithIndexPassedStringLength_ThrowsArgumentOutOfException()
    On Error Resume Next
    
    Char.IsHighSurrogateStr "a", 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsLowSurrogateStr_WithValues_ReturnsExpected(ByRef s As String, ByVal Index As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.IsLowSurrogateStr(s, Index)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLowSurrogateStr_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.IsLowSurrogateStr "a", -1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub IsLowSurrogateStr_WithIndexPassedStringLength_ThrowsArgumentOutOfException()
    On Error Resume Next
    
    Char.IsLowSurrogateStr "a", 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByVal a As Long, ByVal b As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = Char.Equals(a, b)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Function Compare_WithValues_ReturnsExpected(ByVal a As Long, ByVal b As Long) As Long
    Compare_WithValues_ReturnsExpected = Char.Compare(a, b)
End Function

Public Sub ConvertToUtf32_WithValues_ReturnsExpected(ByVal High As Long, ByVal Low As Long, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = Char.ConvertToUtf32(High, Low)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ConvertToUtf32_WithInvalidHighSurrogate_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.ConvertToUtf32 0, &HDC00
    
    AssertArgumentOutOfRangeException Err, "HighSurrogate"
End Sub

Public Sub ConvertToUtf32_WithInvalidLowSurrogate_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Char.ConvertToUtf32 &HD800, 0
    
    AssertArgumentOutOfRangeException Err, "LowSurrogate"
End Sub

Public Sub ConvertToUtf32Str_WithValues_ReturnsExpected(ByVal s As String, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = Char.ConvertToUtf32Str(s, 0)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ConvertToUtf32Str_WithHighSurrogateFollowedByNonLowSurrogate_ThrowsArgumentException()
    On Error Resume Next
    
    Char.ConvertToUtf32Str t("\xD800\x0040"), 0
    
    AssertArgumentException Err, "s"
End Sub

Public Sub ConvertToUtf32Str_WithFirstCharAsLowSurrogate_ThrowsArgumentException()
     On Error Resume Next
     
     Char.ConvertToUtf32Str t("\xDC00\xDC00"), 0
     
     AssertArgumentException Err, "s"
End Sub

