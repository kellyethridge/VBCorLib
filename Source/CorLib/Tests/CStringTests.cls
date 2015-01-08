VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToCharArray_WithString_ReturnsArrayOfCharacters"
            Test.Use "Hello", NewIntegers(Asc("H"), Asc("e"), Asc("l"), Asc("l"), Asc("o"))
            Test.Use "", NewIntegers()
            
        Case "FromCharArray_WithArrayOfCharacters_ReturnsExpected"
            Test.Use NewIntegers(Asc("H"), Asc("e"), Asc("l"), Asc("l"), Asc("o")), "Hello"
            Test.Use NewIntegers(), ""
            
    End Select
End Sub

Public Sub Format_WithOnePlaceHolder_ReturnsPlaceHolderReplacedWithParameter()
    Dim Actual As String
    
    Actual = CString.Format("hel{0}o", "l")
    
    Assert.That Actual, Iz.EqualTo("hello")
End Sub

Public Sub Format_WithOneDuplicatePlaceHolder_ReplacesBothPlaceHoldersWithParameter()
    Dim Actual As String
    
    Actual = CString.Format("he{0}{0}o", "l")
    
    Assert.That Actual, Iz.EqualTo("hello")
End Sub

Public Sub ToCharArray_WithString_ReturnsArrayOfCharacters(ByVal StringToConvert As String, ByRef Expected() As Integer)
    Dim Actual() As Integer
    
    Actual = CString.ToCharArray(StringToConvert)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FromCharArray_WithArrayOfCharacters_ReturnsExpected(ByRef CharactersToConvert() As Integer, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.FromCharArray(CharactersToConvert)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub PadLeft_UsingDefaultPaddingChar_ReturnsStringPaddedWithSpaces()
    Dim Actual As String
    
    Actual = CString.PadLeft("abc", 6)
    
    Assert.That Actual, Iz.EqualTo("   abc")
End Sub

Public Sub PadLeft_UsingPeriodPaddingChar_ReturnsStringPaddedWithPeriods()
    Dim Actual As String
    
    Actual = CString.PadLeft("abc", 6, ".")
    
    
    Assert.That Actual, Iz.EqualTo("...abc")
End Sub

Public Sub PadLeft_WithTotalWidthLessThanLengthStringToPad_ReturnsOriginalString()
    Dim Actual As String
    
    Actual = CString.PadLeft("abc", 2)
    
    Assert.That Actual, Iz.EqualTo("abc")
End Sub

Public Sub PadRight_UsingDefaultPaddingChar_ReturnsStringPaddedWithSpaces()
    Dim Actual As String
    
    Actual = CString.PadRight("abc", 6)
    
    Assert.That Actual, Iz.EqualTo("abc   ")
End Sub

Public Sub PadRight_UsingPeriodPaddingChar_ReturnsStringPaddedWithPeriods()
    Dim Actual As String
    
    Actual = CString.PadRight("abc", 6, ".")
    
    Assert.That Actual, Iz.EqualTo("abc...")
End Sub

Public Sub PadRight_WithTotalWidthLessThanLengthOfStringToPad_ReturnsOriginalString()
    Dim Actual As String
    
    Actual = CString.PadRight("abc", 2)
    
    Assert.That Actual, Iz.EqualTo("abc")
End Sub

Public Sub TestTrimStart()
    Dim Chars() As Integer
    Dim s As String
    
    Chars = CString.ToCharArray("abc")
    s = "cabhi"
    s = CString.TrimStart(s, Chars)
    Assert.AreEqual "hi", s
End Sub

Public Sub TestTrimStartNoTrim()
    Dim Chars(0) As Integer
    Dim s As String
    
    Chars(0) = 48
    s = "hello"
    s = CString.TrimStart(s, Chars)
    Assert.AreEqual "hello", s
End Sub

Public Sub TestTrimStartAllTrimmed()
    Dim Chars(0) As Integer
    Dim s As String
    
    Chars(0) = 48
    s = "00000"
    s = CString.TrimStart(s, Chars)
    Assert.AreEqual "", s
End Sub

Public Sub TestTrimEnd()
    Dim Chars() As Integer
    Dim s As String
    
    Chars = CString.ToCharArray("abc")
    s = "hicab"
    s = CString.TrimEnd(s, Chars)
    Assert.AreEqual "hi", s
End Sub

Public Sub TestTrimEndNoTrim()
    Dim Chars(0) As Integer
    Dim s As String
    
    Chars(0) = 48
    s = "hello"
    s = CString.TrimEnd(s, Chars)
    Assert.AreEqual "hello", s
End Sub

Public Sub TestTrimEndAllTrimmed()
    Dim Chars(0) As Integer
    Dim s As String
    
    Chars(0) = 48
    s = "00000"
    s = CString.TrimEnd(s, Chars)
    Assert.AreEqual "", s
End Sub

Public Sub TestTrim()
    Dim s As String
    
    s = "*;|@123***456@|;*"
    s = CString.Trim(s, "*;|@")
    Assert.AreEqual "123***456", s
End Sub

Public Sub TestTrimWhiteSpace()
    Dim s As String
    
    s = vbCr & vbCr & " hi " & vbLf & vbLf
    s = CString.Trim(s)
    Assert.AreEqual "hi", s
End Sub

Public Sub TestTrimSpaces()
    Dim s As String
    
    s = "     "
    s = CString.Trim(s)
    Assert.AreEqual "", s
End Sub

Public Sub TestTrimEmptyString()
    Dim s As String
    s = CString.Trim(s)
    Assert.AreEqual "", s
End Sub

Public Sub TestJoinAll()
    AssertJoin "one two three"
End Sub

Public Sub TestJoinAllCommaDelimiter()
    AssertJoin "one,two,three", ","
End Sub

Public Sub TestJoinFirst2()
    AssertJoin "one two", , , 2
End Sub

Public Sub TestJoinLast2()
    AssertJoin "two,three", ",", 0, 2
End Sub

Public Sub TestJoinCountTooHigh()
    On Error GoTo errTrap
    AssertJoin "", , , 4
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestJoinVariants()
    Dim v(2) As Variant
    
    v(0) = "hi"
    v(1) = "bye"
    v(2) = "again"
    
    Assert.AreEqual "bye,again", CString.Join(v, ",", 1)
End Sub

Public Sub TestRemove()
    Dim s As String
    
    s = "1234567890"
    Assert.AreEqual "1234890", CString.Remove(s, 4, 3)
End Sub

Public Sub TestRemoveBeginning()
    Dim s As String
    
    s = "1234567890"
    Assert.AreEqual "890", CString.Remove(s, 0, 7)
End Sub

Public Sub TestRemoveEnd()
    Dim s As String
    
    s = "1234567890"
    Assert.AreEqual "123", CString.Remove(s, 3, 7)
End Sub

Public Sub TestIndexOfAnyString()
    Assert.AreEqual 1, CString.IndexOfAny("abc", "b")
End Sub

Public Sub TestIndexOfAnyChars()
    Assert.AreEqual 0, CString.IndexOfAny("abc", NewArray(vbInteger, 98, 97))
End Sub

Public Sub TestIndexOfAnyNotFound()
    Assert.AreEqual -1, CString.IndexOfAny("abc", "123")
End Sub

Public Sub TestIndexOfAnySubString()
    Assert.AreEqual 1, CString.IndexOfAny("abc", "b", 1, 1)
End Sub

Public Sub TestLastIndexOfAnyString()
    Assert.AreEqual 2, CString.LastIndexOfAny("abc", "bc")
End Sub

Public Sub TestLastIndexOfAnyNone()
    Assert.AreEqual -1, CString.LastIndexOfAny("abc", "123")
End Sub

Public Sub TestInsertInMiddle()
    Assert.AreEqual "abc123def", CString.Insert("abcdef", 3, "123")
End Sub

Public Sub TestInsertAtBeginning()
    Assert.AreEqual "123abcdef", CString.Insert("abcdef", 0, "123")
End Sub

Public Sub TestInsertAtEnd()
    Assert.AreEqual "abcdef123", CString.Insert("abcdef", 6, "123")
End Sub

Public Sub TestInsertInvalidIndexBefore()
    On Error GoTo errTrap
    Call CString.Insert("abc", -1, "123")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentOutOfRangeException, "Wrong exception type caught."
End Sub

Public Sub TestInsertInvalidIndexAfter()
    On Error GoTo errTrap
    Call CString.Insert("abc", 4, "123")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentOutOfRangeException, "Wrong exception type caught."
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AssertJoin(ByVal expectedvalue As String, Optional ByVal Delimiter As Variant, Optional ByVal Index As Variant, Optional ByVal Count As Variant)
    Dim s(-1 To 1) As String
    
    s(-1) = "one"
    s(0) = "two"
    s(1) = "three"
    
    Assert.AreEqual expectedvalue, CString.Join(s, Delimiter, Index, Count)
End Sub


