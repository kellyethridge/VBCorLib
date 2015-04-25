VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Const CHARACTER_ZERO As Long = 48

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Append_WithValues_Appends"
            Test.Use "a", "a"
            Test.Use 1, "1"
            Test.Use 1.2, "1.2"
            Test.Use True, "True"
            Test.Use String$(10000, "a"), String$(10000, "a")
            Test.Use "", ""
            Test.Use MakeInt32(5), "5"
            
        Case "Replace_WithValues_Replaces"
            Test.Use "abcd", "bc", "45", "a45d"
            Test.Use "abcd", "ab", "45", "45cd"
            Test.Use "abcd", "cd", "45", "ab45"
            Test.Use "abcd", "a", "aa", "aabcd"
            Test.Use "abcd", "d", "aa", "abcaa"
            Test.Use "abcd", "b", "aa", "aaacd"
            Test.Use "abcd", "bc", "a", "aad"
            Test.Use "abcd", "bc", "", "ad"
            Test.Use "1111", "1", "aa", "aaaaaaaa"
            Test.Use "1212", "1", "aa", "aa2aa2"
            Test.Use "1111", "11", "a", "aa"
            Test.Use "abbbbc", "bd", "def", "abbbbc"
            Test.Use "hahahahahahahahahaha", "ha", "blahblah", "blahblahblahblahblahblahblahblahblahblahblahblahblahblahblahblahblahblahblahblah"
            Test.Use "blahblahblahblahblahblahblahblahblahblahblahblahblahblahblahblahblahblahblahblah", "blahblah", "ha", "hahahahahahahahahaha"
            Test.Use "abcabcabc", "abcf", "", "abcabcabc"
    End Select
End Sub

Public Sub Append_WithValues_Appends(ByRef Value As Variant, ByVal Expected As String)
    Dim sb As New StringBuilder
    
    sb.Append Value
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Append_WithMultipleCalls_Appends()
    Dim sb As New StringBuilder
    sb.Append "a"
    
    sb.Append "b"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("ab")
End Sub

Public Sub Append_WithHugeMultipleCalls_Appends()
    Dim sb As New StringBuilder
    sb.Append Space$(10000)
    
    sb.Append Space$(10000)
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo(Space$(20000))
End Sub

Public Sub Append_WithStartIndex_AppendSubString()
    Dim sb As New StringBuilder
    
    sb.Append "123456", 3
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("456")
End Sub

Public Sub Append_WithStartIndexAndCount_AppendsSubString()
    Dim sb As New StringBuilder
    
    sb.Append "123456", 2, 2
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("34")
End Sub

Public Sub Append_WithStartIndexLessThanZero_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.Append "abc", -1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub Append_WithCountLessThanZero_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.Append "abc", , -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Append_WithStartIndexAndCountPassedLengthOfValue_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.Append "abc", 0, 4
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub AppendChar_WhenCalled_Appends()
    Dim sb As New StringBuilder
    
    sb.AppendChar CHARACTER_ZERO, 50
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo(String$(50, "0"))
End Sub

Public Sub AppendChar_WithDefaultCount_AppendsOneCharacter()
    Dim sb As New StringBuilder
    
    sb.AppendChar CHARACTER_ZERO
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("0")
End Sub

Public Sub AppendChar_WithCountOfZero_DoesNotAppend()
    Dim sb As New StringBuilder
    
    sb.AppendChar CHARACTER_ZERO, 0
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.Empty
End Sub

Public Sub AppendChar_WithCountLessThanZero_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.AppendChar CHARACTER_ZERO, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Chars_WithGet_ReturnsCharacter()
    Dim sb As New StringBuilder
    sb.Append "abc"
    
    Dim Actual As Long
    Actual = sb(1)
    
    Assert.That Actual, Iz.EqualTo(Asc("b"))
End Sub

Public Sub Chars_WithSet_ReplacesCharacter()
    Dim sb As New StringBuilder
    sb.Append "1111"
    
    sb(2) = CHARACTER_ZERO
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("1101")
End Sub

Public Sub AppendChars_WithEmptyArray_DoesNotAppend()
    Dim sb As New StringBuilder
    Dim Chars() As Integer
    Chars = NewIntegers()
    
    sb.AppendChars Chars
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.Empty
End Sub

Public Sub AppendChars_WithNullArray_Throws()
    Dim sb As New StringBuilder
    Dim Chars() As Integer
    
    On Error Resume Next
    sb.AppendChars Chars
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub AppendChars_WithCharacters_Appends()
    Dim sb As New StringBuilder
    Dim Chars() As Integer
    Chars = NewIntegers(49, 50, 51)
    
    sb.AppendChars Chars
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("123")
End Sub

Public Sub AppendChars_WithStartIndexLessThanLBound_Throws()
    Dim sb As New StringBuilder
    Dim Chars() As Integer
    Chars = NewIntegers(49)
    
    On Error Resume Next
    sb.AppendChars Chars, -1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub AppendChars_WithStartIndexAndCharCountLongerThanChars_Throws()
    Dim sb As New StringBuilder
    Dim Chars() As Integer
    Chars = NewIntegers(49, 50, 51)
    
    On Error Resume Next
    sb.AppendChars Chars, 0, 4
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub AppendChars_WithMultiDimensionArray_Throws()
    Dim sb As New StringBuilder
    Dim Chars(0, 0) As Integer
    
    On Error Resume Next
    sb.AppendChars Chars
    
    AssertRankException Err
End Sub

Public Sub AppendChars_WithStartIndexAndCharCount_AppendsSubArray()
    Dim sb As New StringBuilder
    Dim Chars() As Integer
    Chars = NewIntegers(65, 66, 67, 68, 69)
    
    sb.AppendChars Chars, 1, 3
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("BCD")
End Sub

Public Sub Length_WhenSetLongerThanInstance_AppendsSpaces()
    Dim sb As New StringBuilder
    sb.Append "12345"
    
    sb.Length = 10
    
    Dim Actual As String
    Actual = sb.ToString
    Const Expected As String = "12345     "
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Length_WhenSetShortThanInstance_Truncates()
    Dim sb As New StringBuilder
    sb.Append "1234567890"
    
    sb.Length = 5
    
    Dim Actual As String
    Actual = sb.ToString
    Const Expected As String = "12345"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Length_WhenLessThanZero_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.Length = -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub Insert_WithZeroIndex_InsertsAtBeginning()
    Dim sb As New StringBuilder
    sb.Append "1234"
    
    sb.Insert 0, "abc"
    
    Dim Actual As String
    Actual = sb.ToString
    Const Expected = "abc1234"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Insert_WithNonZeroIndex_InsertsBetweenCharacters()
    Dim sb As New StringBuilder
    sb.Append "1234"
    
    sb.Insert 2, "abc"
    
    Dim Actual As String
    Actual = sb.ToString
    Const Expected = "12abc34"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Insert_WithIndexAtEnd_AppendsToInstance()
    Dim sb As New StringBuilder
    sb.Append "1234"
    
    sb.Insert 4, "abc"
    
    Dim Actual As String
    Actual = sb.ToString
    Const Expected = "1234abc"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Insert_WithIndexLessThanZero_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.Insert -1, "abc"
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub Insert_WithCountLessThanZero_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.Insert 0, "abc", -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Insert_WithCountOfZero_DoesNotInsert()
    Dim sb As New StringBuilder
    sb.Append "abc"
    
    sb.Insert 0, "abc", 0
    
    Dim Actual As String
    Actual = sb.ToString
    Const Expected As String = "abc"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Insert_WithMultipleCounts_InsertsMultipleTimes()
    Dim sb As New StringBuilder
    sb.Append "1234"
    
    sb.Insert 2, "abc", 2
    
    Dim Actual As String
    Actual = sb.ToString
    Const Expected As String = "12abcabc34"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub InsertChars_WithNullArray_Throws()
    Dim sb As New StringBuilder
    Dim Chars() As Integer
    
    On Error Resume Next
    sb.InsertChars 0, Chars
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub InsertChars_WithEmptyArray_InsertsNothing()
    Dim sb As New StringBuilder
    sb.Append "1234"
    Dim Chars() As Integer
    Chars = NewIntegers
    
    sb.InsertChars 2, Chars
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("1234")
End Sub

Public Sub InsertChars_WithIndexTo3rdPosition_InsertsInBetween()
    Dim sb As New StringBuilder
    sb.Append "1234"
    Dim Chars() As Integer
    Chars = NewIntegers(97, 97, 97)
    
    sb.InsertChars 2, Chars
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("12aaa34")
End Sub

Public Sub InsertChars_WithIndexOfZero_InsertsAtBeginning()
    Dim sb As New StringBuilder
    sb.Append "1234"
    Dim Chars() As Integer
    Chars = NewIntegers(97, 97, 97)
    
    sb.InsertChars 0, Chars
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("aaa1234")
End Sub

Public Sub InsertChars_WithIndexOfLengthOfInstance_InsertsAtEnd()
    Dim sb As New StringBuilder
    sb.Append "1234"
    Dim Chars() As Integer
    Chars = NewIntegers(97, 97, 97)
    
    sb.InsertChars 4, Chars
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("1234aaa")
End Sub

Public Sub InsertChars_WithSubArray_InsertsPortion()
    Dim sb As New StringBuilder
    sb.Append "1234"
    Dim Chars() As Integer
    Chars = NewIntegers(97, 98, 99, 100)
    
    sb.InsertChars 2, Chars, 1, 2
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("12bc34")
End Sub

Public Sub InsertChars_WithIndexLessThanZero_Throws()
    Dim sb As New StringBuilder
    Dim Chars() As Integer
    ReDim Chars(1 To 1)
    
    On Error Resume Next
    sb.InsertChars -1, Chars
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub InsertChars_WithStartIndexLessThanLBound_Throws()
    Dim sb As New StringBuilder
    Dim Chars() As Integer
    ReDim Chars(1 To 1)
    
    On Error Resume Next
    sb.InsertChars 0, Chars, 0
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub InsertChars_WithNegativeCount_Throws()
    Dim sb As New StringBuilder
    Dim Chars(0) As Integer
    
    On Error Resume Next
    sb.InsertChars 0, Chars, 0, -1
    
    AssertArgumentOutOfRangeException Err, "CharCount"
End Sub

Public Sub InsertChars_WithStartIndexAndCountExtendingPassedUBound_Throws()
    Dim sb As New StringBuilder
    Dim Chars(1 To 3) As Integer
    
    On Error Resume Next
    sb.InsertChars 0, Chars, 1, 4
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub Remove_WithStartIndexLessThanZero_Throws()
    Dim sb As New StringBuilder
    sb.Append "a"
    
    On Error Resume Next
    sb.Remove -1, 1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub Remove_WithCountLessThanZero_Throws()
    Dim sb As New StringBuilder
    sb.Append "a"
    
    On Error Resume Next
    sb.Remove 0, -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub Remove_WithStartIndexAndCountGreaterThanInstanceLength_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    
    On Error Resume Next
    sb.Remove 0, 4
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub Remove_WithStartIndexZero_RemovesFromBeginning()
    Dim sb As New StringBuilder
    sb.Append "1234567890"
    
    sb.Remove 0, 5
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("67890")
End Sub

Public Sub Remove_WithStartIndexAndLengthCoverEnd_RemovesFromEnd()
    Dim sb As New StringBuilder
    sb.Append "1234567890"
    
    sb.Remove 5, 5
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("12345")
End Sub

Public Sub Replace_WithEmptyNewValue_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    
    On Error Resume Next
    sb.Replace "", "a"
    
    AssertArgumentException Err, "OldValue"
End Sub

Public Sub Replace_WithStartIndexLessThanZero_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    
    On Error Resume Next
    sb.Replace "a", "b", -1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub Replace_WithCountLessThanZero_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    
    On Error Resume Next
    sb.Replace "a", "b", 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Replace_WithStartIndexAndCountGreaterThanInstanceLength_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    
    On Error Resume Next
    sb.Replace "a", "b", 0, 4
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub Replace_WithValues_Replaces(ByVal Initial As String, ByVal OldValue As String, ByVal NewValue As String, ByVal Expected As String)
    Dim sb As New StringBuilder
    sb.Append Initial
    
    sb.Replace OldValue, NewValue
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Replace_WithinSubstring_ReplacesSubInstancesOnly()
    Dim sb As New StringBuilder
    sb.Append "111111"
    
    sb.Replace "11", "a", 2, 2
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("11a11")
End Sub

Public Sub Replace_WithSubstringStartingInMiddleOfMatch_DoesNotReplacePartialMatch()
    Dim sb As New StringBuilder
    sb.Append "blahblahblahblahblahblahblahblah"
    
    sb.Replace "blah", "ha", 2
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("blahhahahahahahaha")
End Sub

Public Sub Replace_UsingXBeatTest()
    Dim sb As New StringBuilder
    sb.Insert 0, "xx" & Space$(100) & "xx", 10
    
    sb.Replace "x", Space$(100)
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo(Space$(5000))
End Sub

Public Sub AppendFormat_WithPlaceHolder_InsertsValue()
    Dim sb As New StringBuilder
    
    sb.AppendFormat "Hello {0}", "World"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("Hello World")
End Sub

Public Sub AppendFormat_WithLengthInfo_IncludeLengthInfo()
    Dim sb As New StringBuilder
    
    sb.AppendFormat "Hello {0, -5}", 1
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("Hello 1    ")
End Sub

Public Sub AppendFormat_WithFormatInfo_IncludeFormatInfo()
    Dim sb As New StringBuilder
    
    sb.AppendFormat "{0:X2}", 10
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("0A")
End Sub

Public Sub AppendFormat_WithAlignmentAndFormatInfo_IncludesBoth()
    Dim sb As New StringBuilder
    
    sb.AppendFormat "{0,5:X2}", 10
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("   0A")
End Sub

Public Sub AppendFormat_WithMultiplePlaceHolders_FormatsEach()
    Dim sb As New StringBuilder
    
    sb.AppendFormat "{0} {1}", "a", "b"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("a b")
End Sub

Public Sub AppendFormat_WithRepeatedPlaceHolder_RepeatsFormat()
    Dim sb As New StringBuilder
    
    sb.AppendFormat "{0} {0}", "a"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("a a")
End Sub

Public Sub AppendFormat_WithUnpairedOpenBrace_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.AppendFormat "{0", "a"
    
    AssertFormatException Err
End Sub

Public Sub AppendFormat_WithUnpairedOpenBrace_Throws2()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.AppendFormat " {0 ", "a"
    
    AssertFormatException Err
End Sub

Public Sub AppendFormat_WithUnpairedClosingBrace_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.AppendFormat "0}", "a"
    
    AssertFormatException Err
End Sub

Public Sub AppendFormat_WithUnpairedClosingBrace_Throws2()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.AppendFormat " 0} ", "a"
    
    AssertFormatException Err
End Sub

Public Sub AppendFormat_WithNegativePlaceHolder_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.AppendFormat "{-1}", "a"
    
    AssertFormatException Err
End Sub

Public Sub AppendFormat_WithMorePlaceHoldersThanValues_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.AppendFormat "{0} {1}", "a"
    
    AssertFormatException Err
End Sub

Public Sub AppendFormat_WithReplicatedBrace_EscapesBrace()
    Dim sb As New StringBuilder
    
    sb.AppendFormat "{{0}}"
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("{0}")
End Sub

Public Sub AppendFormat_WithEscapedOpeningBraceAndNonEscapedClosingBrace_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.AppendFormat " {{0} ", "hi"
    
    AssertFormatException Err
End Sub

Public Sub AppendFormat_WithNonEscapedOpeningBraceAndEscapedClosingBrace_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.AppendFormat " {0}} ", "hi"
    
    AssertFormatException Err
End Sub

Public Sub AppendFormat_PlaceHolderWithMissingIndex_Throws()
    Dim sb As New StringBuilder
    
    On Error Resume Next
    sb.AppendFormat "{ }", "hi"
    
    AssertFormatException Err
End Sub

Public Sub AppendFormatEx_WithCustomFormatter_RequestsCorrectProvider()
    Dim sb As New StringBuilder
    Dim f As New FakeCustomFormatter
    f.Disable = True
    
    sb.AppendFormatEx f, "{0}", "a"
    
    Assert.That f.FormatType, Iz.EqualTo("ICustomFormatter").IgnoreCase
End Sub

Public Sub AppendFormatEx_WithNonCustomFormatter_DefaultFormats()
    Dim sb As New StringBuilder
    Dim f As New FakeCustomFormatter
    f.Disable = True
    
    sb.AppendFormatEx f, "{0:X2}", 10
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("0A")
End Sub

Public Sub AppendFormatEx_WithFormatter_ReturnsFormatted()
    Dim sb As New StringBuilder
    Dim f As New FakeCustomFormatter
    f.FormatResult = "abc"
    
    sb.AppendFormatEx f, "{0:X2}", 10
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("abc")
    Assert.That f.FormatArg, Iz.EqualTo(10)
    Assert.That f.FormatText, Iz.EqualTo("X2")
End Sub

Public Sub AppendFormatEx_WithNullReturn_PerformsDefaultFormatting()
    Dim sb As New StringBuilder
    Dim f As New FakeCustomFormatter
    f.ReturnNull = True
    
    sb.AppendFormatEx f, "{0:X2}", 10
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.EqualTo("0A")
End Sub

Public Sub AppendFormatEx_WithEmptyReturn_FormatsAsEmpty()
    Dim sb As New StringBuilder
    Dim f As New FakeCustomFormatter
    f.FormatResult = ""
    
    sb.AppendFormatEx f, "{0:x2}", 10
    
    Dim Actual As String
    Actual = sb.ToString
    Assert.That Actual, Iz.Empty
End Sub

Public Sub CopyTo_WhenCalled_CopiesStringToArray()
    Dim sb As New StringBuilder
    sb.Append "0123456789"
    Dim Chars() As Integer
    ReDim Chars(2)
    
    sb.CopyTo 2, Chars, 0, 3
    
    Dim Expected() As Integer
    Expected = NewIntegers(50, 51, 52)
    Assert.That Chars, Iz.EqualTo(Expected)
End Sub

Public Sub CopyTo_WithNullDestination_Throw()
    Dim sb As New StringBuilder
    sb.Append "abc"
    Dim Chars() As Integer
    
    On Error Resume Next
    sb.CopyTo 0, Chars, 0, 3
    
    AssertArgumentNullException Err, "Destination"
End Sub

Public Sub CopyTo_WithNegativeSourceIndex_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    Dim Chars(0) As Integer
    
    On Error Resume Next
    sb.CopyTo -1, Chars, 0, 1
    
    AssertArgumentOutOfRangeException Err, "SourceIndex"
End Sub

Public Sub CopyTo_WithSourceIndexLongerThanInstance_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    Dim Chars(0) As Integer
    
    On Error Resume Next
    sb.CopyTo 4, Chars, 0, 1
    
    AssertArgumentOutOfRangeException Err, "SourceIndex"
End Sub

Public Sub CopyTo_WithDestinationIndexLessThanLBound_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    Dim Chars(5 To 6) As Integer
    
    On Error Resume Next
    sb.CopyTo 0, Chars, 4, 2
    
    AssertArgumentOutOfRangeException Err, "DestinationIndex"
End Sub

Public Sub CopyTo_WithSourceIndexAndCountGreaterThanTheInstance_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    Dim Chars(0 To 6) As Integer
    
    On Error Resume Next
    sb.CopyTo 0, Chars, 0, 4
    
    AssertArgumentException Err
End Sub

Public Sub CopyTo_WithDestinationIndexAndCountGreaterThanUBound_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    Dim Chars(0 To 6) As Integer
    
    On Error Resume Next
    sb.CopyTo 0, Chars, 5, 3
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub ToString_WithNegativeStartIndex_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    
    On Error Resume Next
    sb.ToString -1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub ToString_WithNegativeLength_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    
    On Error Resume Next
    sb.ToString 0, -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub ToString_WithStartIndexAndLengthGreaterThanInstance_Throws()
    Dim sb As New StringBuilder
    sb.Append "abc"
    
    On Error Resume Next
    sb.ToString 0, 4
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub ToString_WithSubString_ReturnsSubPortion()
    Dim sb As New StringBuilder
    sb.Append "abcd"
    
    Dim Actual As String
    Actual = sb.ToString(2, 2)
    
    Assert.That Actual, Iz.EqualTo("cd")
End Sub

