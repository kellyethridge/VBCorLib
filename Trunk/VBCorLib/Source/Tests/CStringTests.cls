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
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToCharArray_WithString_ReturnsArrayOfCharacters"
            Test.Use "Hello", NewIntegers(Asc("H"), Asc("e"), Asc("l"), Asc("l"), Asc("o"))
            Test.Use "", NewIntegers()
            
        Case "FromCharArray_WithArrayOfCharacters_ReturnsExpected"
            Test.Use NewIntegers(Asc("H"), Asc("e"), Asc("l"), Asc("l"), Asc("o")), "Hello"
            Test.Use NewIntegers(), ""
        
        Case "IsEmptyOrWhitespace_WithValues_ReturnsExpected"
            Test.Use "", True
            Test.Use "a", False
            Test.Use " ", True
            Test.Use " a ", False
            Test.Use WhiteSpaces, True
            
        Case "TrimStart_WithDefaultWhiteSpace_TrimsWhiteSpaceFromStart"
            Test.Use "", ""
            Test.Use WhiteSpaces, ""
            Test.Use "hello" & WhiteSpaces, "hello" & WhiteSpaces
            Test.Use WhiteSpaces & "hello", "hello"
        
        Case "TrimStart_WithValues_ReturnsExpected"
            Test.Use "", WhiteSpaces, ""
            Test.Use WhiteSpaces & "hello" & WhiteSpaces, WhiteSpaces, "hello" & WhiteSpaces
            Test.Use "hello" & WhiteSpaces, "helo", WhiteSpaces
            Test.Use "hello" & WhiteSpaces, NewIntegers(104, 101, 108, 111), WhiteSpaces
            Test.Use "hello", "helo", ""
            Test.Use "hello", NewIntegers(104, 101, 108, 111), ""
            Test.Use "hello", WhiteSpaces, "hello"
            Test.Use WhiteSpaces & "hello", "", WhiteSpaces & "hello"
            
        Case "TrimEnd_WithDefaultWhiteSpace_TrimsWhiteSpaceFromEnd"
            Test.Use "", ""
            Test.Use WhiteSpaces, ""
            Test.Use "hello" & WhiteSpaces, "hello"
            Test.Use WhiteSpaces & "hello", WhiteSpaces & "hello"
        
        Case "TrimEnd_WithValues_ReturnsExpected"
            Test.Use "", WhiteSpaces, ""
            Test.Use WhiteSpaces & "hello" & WhiteSpaces, WhiteSpaces, WhiteSpaces & "hello"
            Test.Use WhiteSpaces & "hello", "helo", WhiteSpaces
            Test.Use WhiteSpaces & "hello", NewIntegers(104, 101, 108, 111), WhiteSpaces
            Test.Use "hello", "helo", ""
            Test.Use "hello", NewIntegers(104, 101, 108, 111), ""
            Test.Use "hello", WhiteSpaces, "hello"
            Test.Use "hello" & WhiteSpaces, "", "hello" & WhiteSpaces
        
        Case "Trim_WithDefaultWhiteSpace_TrimsWhiteSpace"
            Test.Use "", ""
            Test.Use WhiteSpaces, ""
            Test.Use "hello" & WhiteSpaces, "hello"
            Test.Use WhiteSpaces & "hello", "hello"
            
        Case "Trim_WithValues_ReturnsExpected"
            Test.Use "", WhiteSpaces, ""
            Test.Use WhiteSpaces & "hello" & WhiteSpaces, WhiteSpaces, "hello"
            Test.Use "hello" & WhiteSpaces, "helo", WhiteSpaces
            Test.Use WhiteSpaces & "hello", "helo", WhiteSpaces
            Test.Use "hello", "helo", ""
            Test.Use "hello", NewIntegers(104, 101, 108, 111), ""
            Test.Use "hello", WhiteSpaces, "hello"
            
    End Select
End Sub

Public Sub IsEmptyOrWhitespace_WithValues_ReturnsExpected(ByVal s As String, ByVal Expected As Boolean)
    Dim Actual As Boolean

    Actual = cString.IsNullOrWhiteSpace(s)

    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Format_WithOnePlaceHolder_ReturnsPlaceHolderReplacedWithParameter()
    Dim Actual As String
    
    Actual = cString.Format("hel{0}o", "l")
    
    Assert.That Actual, Iz.EqualTo("hello")
End Sub

Public Sub Format_WithOneDuplicatePlaceHolder_ReplacesBothPlaceHoldersWithParameter()
    Dim Actual As String
    
    Actual = cString.Format("he{0}{0}o", "l")
    
    Assert.That Actual, Iz.EqualTo("hello")
End Sub

Public Sub ToCharArray_WithString_ReturnsArrayOfCharacters(ByVal StringToConvert As String, ByRef Expected() As Integer)
    Dim Actual() As Integer
    
    Actual = cString.ToCharArray(StringToConvert)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToCharArray_WithSubstringSpecified_ReturnsSubPortion()
    Dim Actual() As Integer
    
    Actual = cString.ToCharArray("hello", 1, 3)
    
    Assert.That Actual, Iz.EqualTo(NewIntegers(101, 108, 108))
End Sub

Public Sub ToCharArray_WithStartIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    cString.ToCharArray "hello", -1, 1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub ToCharArray_WithLengthLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    cString.ToCharArray "hello", 0, -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub ToCharArray_WithStartIndexAndLengthGreaterThanTheSizeOfCharsArray_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    cString.ToCharArray "hello", 1, 5
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub FromCharArray_WithArrayOfCharacters_ReturnsExpected(ByRef CharactersToConvert() As Integer, ByVal Expected As String)
    Dim Actual As String
    
    Actual = cString.FromCharArray(CharactersToConvert)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FromCharArray_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Chars(1 To 10) As Integer
    
    On Error Resume Next
    cString.FromCharArray Chars, 0, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub FromCharArray_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Chars(0) As Integer
    
    On Error Resume Next
    cString.FromCharArray Chars, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub FromCharArray_WithIndexAndCountBeyondUBound_ThrowsArgumentExceptin()
    Dim Chars(0 To 5) As Integer
    
    On Error Resume Next
    cString.FromCharArray Chars, 1, 6
    
    AssertArgumentException Err
End Sub

Public Sub FromCharArray_WithNullArray_ThrowsArgumentNullException()
    Dim NullChars() As Integer
    
    On Error Resume Next
    cString.FromCharArray NullChars
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub FromCharArray_WithMultiDimArray_ThrowsRankException()
    Dim MultiDim(0, 0) As Integer
    
    On Error Resume Next
    cString.FromCharArray MultiDim
    
    AssertRankException Err
End Sub

Public Sub PadLeft_UsingDefaultPaddingChar_ReturnsStringPaddedWithSpaces()
    Dim Actual As String
    
    Actual = cString.PadLeft("abc", 6)
    
    Assert.That Actual, Iz.EqualTo("   abc")
End Sub

Public Sub PadLeft_UsingPeriodPaddingChar_ReturnsStringPaddedWithPeriods()
    Dim Actual As String
    
    Actual = cString.PadLeft("abc", 6, ".")
        
    Assert.That Actual, Iz.EqualTo("...abc")
End Sub

Public Sub PadLeft_WithTotalWidthLessThanLengthStringToPad_ReturnsOriginalString()
    Dim Actual As String
    
    Actual = cString.PadLeft("abc", 2)
    
    Assert.That Actual, Iz.EqualTo("abc")
End Sub

Public Sub PadRight_UsingDefaultPaddingChar_ReturnsStringPaddedWithSpaces()
    Dim Actual As String
    
    Actual = cString.PadRight("abc", 6)
    
    Assert.That Actual, Iz.EqualTo("abc   ")
End Sub

Public Sub PadRight_UsingPeriodPaddingChar_ReturnsStringPaddedWithPeriods()
    Dim Actual As String
    
    Actual = cString.PadRight("abc", 6, ".")
    
    Assert.That Actual, Iz.EqualTo("abc...")
End Sub

Public Sub PadRight_WithTotalWidthLessThanLengthOfStringToPad_ReturnsOriginalString()
    Dim Actual As String
    
    Actual = cString.PadRight("abc", 2)
    
    Assert.That Actual, Iz.EqualTo("abc")
End Sub

Public Sub TrimStart_WithDefaultWhiteSpace_TrimsWhiteSpaceFromStart(ByVal s As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = cString.TrimStart(s)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TrimStart_WithValues_ReturnsExpected(ByVal s As String, ByRef TrimChars As Variant, ByVal Expected As String)
    Dim Actual As String
    
    Actual = cString.TrimStart(s, TrimChars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TrimEnd_WithDefaultWhiteSpace_TrimsWhiteSpaceFromEnd(ByVal s As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = cString.TrimEnd(s)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TrimEnd_WithValues_ReturnsExpected(ByVal s As String, ByRef TrimChars As Variant, ByVal Expected As String)
    Dim Actual As String
    
    Actual = cString.TrimEnd(s, TrimChars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Trim_WithDefaultWhiteSpace_TrimsWhiteSpace(ByVal s As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = cString.Trim(s)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Trim_WithValues_ReturnsExpected(ByVal s As String, ByRef TrimChars As Variant, ByVal Expected As String)
    Dim Actual As String
    
    Actual = cString.Trim(s, TrimChars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Join_WhenCalled_ReturnsDelimitedString()
    Dim Value() As String
    Value = NewStrings("a", "b", "c")
    
    Dim Actual As String
    Actual = cString.Join(",", Value)
    
    Assert.That Actual, Iz.EqualTo("a,b,c")
End Sub

Public Sub Join_WithSubset_ReturnsDelimitedSubset()
    Dim Value() As String
    Value = NewStrings("a", "b", "c", "d")
    
    Dim Actual As String
    Actual = cString.Join(",", Value, 1, 2)
    
    Assert.That Actual, Iz.EqualTo("b,c")
End Sub

Public Sub Join_WithStartIndexBelowLBound_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    cString.Join ",", Array(1, 1), -1, 2
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub Join_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    cString.Join ",", Array(1, 1), 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Join_WithStartIndexPlusCountIsInvalid_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    cString.Join ",", Array(1, 2), 0, 3
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub Join_WithLongs_ReturnsLongsJoined()
    Dim Actual As String
    
    Actual = cString.Join(",", Array(1, 2, 3, 4))
    
    Assert.That Actual, Iz.EqualTo("1,2,3,4")
End Sub

Public Sub Join_WithIObjects_ReturnsJoined()
    Dim Objs(0 To 1) As Variant
    Set Objs(0) = NewInt32(32)
    Set Objs(1) = NewInt32(1)
    
    Dim Actual As String
    Actual = cString.Join(",", Objs)
    
    Assert.That Actual, Iz.EqualTo("32,1")
End Sub

'Public Sub TestRemove()
'    Dim s As String
'
'    s = "1234567890"
'    Assert.AreEqual "1234890", cString.Remove(s, 4, 3)
'End Sub
'
'Public Sub TestRemoveBeginning()
'    Dim s As String
'
'    s = "1234567890"
'    Assert.AreEqual "890", cString.Remove(s, 0, 7)
'End Sub
'
'Public Sub TestRemoveEnd()
'    Dim s As String
'
'    s = "1234567890"
'    Assert.AreEqual "123", cString.Remove(s, 3, 7)
'End Sub
'
'Public Sub TestIndexOfAnyString()
'    Assert.AreEqual 1, cString.IndexOfAny("abc", "b")
'End Sub
'
'Public Sub TestIndexOfAnyChars()
'    Assert.AreEqual 0, cString.IndexOfAny("abc", NewArray(vbInteger, 98, 97))
'End Sub
'
'Public Sub TestIndexOfAnyNotFound()
'    Assert.AreEqual -1, cString.IndexOfAny("abc", "123")
'End Sub
'
'Public Sub TestIndexOfAnySubString()
'    Assert.AreEqual 1, cString.IndexOfAny("abc", "b", 1, 1)
'End Sub
'
'Public Sub TestLastIndexOfAnyString()
'    Assert.AreEqual 2, cString.LastIndexOfAny("abc", "bc")
'End Sub
'
'Public Sub TestLastIndexOfAnyNone()
'    Assert.AreEqual -1, cString.LastIndexOfAny("abc", "123")
'End Sub
'
'Public Sub TestInsertInMiddle()
'    Assert.AreEqual "abc123def", cString.Insert("abcdef", 3, "123")
'End Sub
'
'Public Sub TestInsertAtBeginning()
'    Assert.AreEqual "123abcdef", cString.Insert("abcdef", 0, "123")
'End Sub
'
'Public Sub TestInsertAtEnd()
'    Assert.AreEqual "abcdef123", cString.Insert("abcdef", 6, "123")
'End Sub
'
'Public Sub TestInsertInvalidIndexBefore()
'    On Error GoTo errTrap
'    Call cString.Insert("abc", -1, "123")
'    On Error GoTo 0
'    Assert.Fail "An exception should be thrown."
'
'errTrap:
'    Dim Ex As Exception
'    Assert.IsTrue Catch(Ex), "An exception should be caught."
'    Assert.IsTrue TypeOf Ex Is ArgumentOutOfRangeException, "Wrong exception type caught."
'End Sub
'
'Public Sub TestInsertInvalidIndexAfter()
'    On Error GoTo errTrap
'    Call cString.Insert("abc", 4, "123")
'    On Error GoTo 0
'    Assert.Fail "An exception should be thrown."
'
'errTrap:
'    Dim Ex As Exception
'    Assert.IsTrue Catch(Ex), "An exception should be caught."
'    Assert.IsTrue TypeOf Ex Is ArgumentOutOfRangeException, "Wrong exception type caught."
'End Sub
'
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''   Private Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Sub AssertJoin(ByVal expectedvalue As String, Optional ByVal Delimiter As Variant, Optional ByVal Index As Variant, Optional ByVal Count As Variant)
'    Dim s(-1 To 1) As String
'
'    s(-1) = "one"
'    s(0) = "two"
'    s(1) = "three"
'
'    Assert.AreEqual expectedvalue, cString.Join(s, Delimiter, Index, Count)
'End Sub


Private Property Get WhiteSpaces() As String
    WhiteSpaces = ChrW$(&H20) & ChrW$(&HD) & ChrW$(&H9) & ChrW$(&HA) & ChrW$(&HB) & ChrW$(&HC) & ChrW$(&H85) & ChrW$(&HA0) & ChrW$(&H1680) & ChrW$(&H180E) & _
        ChrW$(&H2000) & ChrW$(&H2001) & ChrW$(&H2002) & ChrW$(&H2003) & ChrW$(&H2004) & ChrW$(&H2005) & ChrW$(&H2006) & ChrW$(&H2007) & ChrW$(&H2008) & ChrW$(&H2009) & ChrW$(&H200A) & _
        ChrW$(&H2028) & ChrW$(&H2029) & ChrW$(&H202F) & ChrW$(&H205F) & _
        ChrW$(&H3000)
End Property

Private Property Get WhiteSpaceChars() As Integer()
    WhiteSpaceChars = NewIntegers(&H20, &HD, &H9, &HA, &HB, &HC, &H85, &HA0, &H1680, &H180E, &H2000, &H2001, &H2002, &H2003, &H2004, &H2005, &H2006, &H2007, &H2008, &H2009, &H200A, &H2028, &H2029, &H202F, &H205F, &H3000)
End Property



