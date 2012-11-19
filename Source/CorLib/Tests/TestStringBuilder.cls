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

Private Const CHARACTER_ZERO As Long = 48

Private sb As StringBuilder


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
            
    End Select
End Sub

Public Sub Setup()
    Set sb = New StringBuilder
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
    
    Assert.AreEqual String$(50, "0"), sb.ToString
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

Public Sub TestSetLengthPastCapacity()
    sb.Capacity = 10
    sb.Append "12345"
    sb.Length = 20
    
    Assert.AreEqual "12345               ", sb.ToString
End Sub

Public Sub TestLength()
    sb.Append "1234567890"
    sb.Length = 5
    sb.Append "abcde"
    Assert.AreEqual "12345abcde", sb.ToString
End Sub

Public Sub TestAppendSubString()
    sb.AppendString "Hello", 2, 3
    Assert.AreEqual "llo", sb.ToString
End Sub



Public Sub TestInsertString()
    sb.Append "1234567890"
    sb.Insert 2, "aaa"
    Assert.AreEqual "12aaa34567890", sb.ToString
End Sub

Public Sub TestInsertBeginning()
    sb.Insert 0, Space$(10), 10
    Assert.AreEqual Space$(100), sb.ToString
End Sub

Public Sub TestInsertAtEnd()
    sb.Append "Hello"
    sb.Insert 5, "Bye"
    Assert.AreEqual "HelloBye", sb.ToString
End Sub

Public Sub TestInsertStringCount()
    sb.Append "1234567890"
    sb.Insert 2, "aaa", 2
    Assert.AreEqual "12aaaaaa34567890", sb.ToString
End Sub

Public Sub TestInsertChars()
    sb.Append "1234567890"
    Dim Chars() As Integer
    Chars = NewIntegers(97, 97, 97)
    sb.Insertchars 2, Chars
    Assert.AreEqual "12aaa34567890", sb.ToString
End Sub

Public Sub TestInsertCharsAtEnd()
    sb.Append "Hello"
    Dim Chars() As Integer
    Chars = NewIntegers(97, 97, 97)
    sb.Insertchars 5, Chars
    Assert.AreEqual "Helloaaa", sb.ToString
End Sub

Public Sub TestInsertSubChars()
    sb.Append "1234567890"
    Dim Chars() As Integer
    Chars = NewIntegers(97, 98, 99, 100)
    sb.Insertchars 2, Chars, 1, 3
    Assert.AreEqual "12bcd34567890", sb.ToString
End Sub

Public Sub TestRemoveFromStart()
    sb.Append "1234567890"
    sb.Remove 0, 5
    Assert.AreEqual "67890", sb.ToString
End Sub

Public Sub TestRemoveFromEnd()
    sb.Append "1234567890"
    sb.Remove 5, 5
    Assert.AreEqual "12345", sb.ToString
End Sub

Public Sub TestReplace()
    sb.Append "1234567890"
    sb.Replace "45", ""
    Assert.AreEqual 8, sb.Length
    Assert.AreEqual "12367890", sb.ToString
End Sub

Public Sub TestReplaceMultiple()
    sb.Insert 0, "Hello", 2
    sb.Replace "ll", "mmmm"
    
    Dim s As String
    Dim i As Long
    For i = 1 To 2
        s = s & "Hemmmmo"
    Next i
    Assert.AreEqual 14, sb.Length
    Assert.AreEqual s, sb.ToString
End Sub

Public Sub TestReplaceWithinSubString()
    sb.Append "\\\\server\\share\\"
    sb.Replace "\\", "\", 2, sb.Length - 2
    
    Assert.AreEqual 16, sb.Length
    Assert.AreEqual "\\\server\share\", sb.ToString
End Sub

Public Sub TestReduce()
    sb.Append "c:\....\abc\"
    
    Dim i As Long
    Do
        i = sb.Length
        Call sb.Replace("...\", "..\")
    Loop While i > sb.Length
    
    Assert.AreEqual "c:\..\abc\", sb.ToString
End Sub

Public Sub TestUnwindReplaceOnPartialMatch()
    sb.Append "abbbbc"
    sb.Replace "bd", ""
    Assert.AreEqual "abbbbc", sb.ToString
End Sub

Public Sub XBeatTest()
    Dim rep As String
    Dim find As String
    Dim src As String
    
    find = "ha"
    rep = "blahblah"
    src = "hahahahahahahahahaha"
    sb.Append src
    sb.Replace find, rep
    Assert.AreEqual Replace$(src, find, rep), sb.ToString
End Sub

Public Sub XBeatTest2()
    sb.Insert 0, "xx" & Space$(10000) & "xx", 10
    sb.Replace "x", Space$(10000)
    Assert.AreEqual Space$(500000), sb.ToString
End Sub

Public Sub TestReplaceLastChar()
    sb.Insert 0, "abc", 10
    sb.Append "d"
    sb.Replace "cd", "1234"
    Assert.AreEqual "abcabcabcabcabcabcabcabcabcab1234", sb.ToString
End Sub

Public Sub TestReplaceAtEnd()
    sb.AppendQuick "MyStuff"
    sb.Replace "Stuff", "Junk"
    Assert.AreEqual "MyJunk", sb.ToString
End Sub

Public Sub TestReplaceInsufficientMatch()
    sb.AppendQuick "MyStuf"
    sb.Replace "Stuff", "Junk"
    Assert.AreEqual "MyStuf", sb.ToString
End Sub

Public Sub TestPartialWordAtEnd()
    sb.AppendQuick "MyStuff"
    sb.Replace "Stuffy", "Junky"
    Assert.AreEqual "MyStuff", sb.ToString
End Sub

Public Sub TestReplaceOfSameSize()
    sb.AppendChar Asc("A"), 10
    sb.Replace "A", "B"
    Assert.AreEqual String$(10, "B"), sb.ToString
End Sub

Public Sub TestAppendFormat()
    sb.AppendFormat "Hello {0}", "World"
    Assert.AreEqual "Hello World", sb.ToString
End Sub

Public Sub TestAppendFormatLeftAdjust()
    sb.AppendFormat "Hello {0, -5 }", "1"
    Assert.AreEqual "Hello 1    ", sb.ToString
End Sub

Public Sub TestAppendFormatRightAdjust()
    sb.AppendFormat "Hello {0, 10 } world!", "kelly"
    Assert.AreEqual "Hello      kelly world!", sb.ToString
End Sub

Public Sub TestAppendFormatMultiArgs()
    sb.AppendFormat "1={0}, 2={1}, 1={0}, Name={2}", 1, 2, "Kelly"
    Assert.AreEqual "1=1, 2=2, 1=1, Name=Kelly", sb.ToString
End Sub

Public Sub TestAppendFormatFormattable()
    Dim i1 As New Int32Formattable
    Dim i2 As New Int32Formattable
    i1.mValue = 10
    i2.mValue = 20
    sb.AppendFormat "10={0}, 20={1}", i1, i2
    Assert.AreEqual "10=10, 20=20", sb.ToString
End Sub

Public Sub TestIObjectBased()
    Dim i1 As New Int32
    Dim i2 As New Int32
    i1.mValue = 10
    i2.mValue = 20
    sb.AppendFormat "10={0}, 20={1}", i1, i2
    Assert.AreEqual "10=10, 20=20", sb.ToString
End Sub

Public Sub TestFormatInfo()
    On Error Resume Next
    Dim i1 As New Int32Formattable
    i1.mValue = 1
    sb.AppendFormat "{0:9}", i1
    Assert.AreEqual "000000001", sb.ToString
End Sub

Public Sub TestAppendFormatNoArgs()
    sb.AppendFormat "Hello"
    Assert.AreEqual "Hello", sb.ToString
End Sub

Public Sub TestAppendFormatInt32()
    sb.AppendFormat "Hello {0:G3}", 1235
    Assert.AreEqual "Hello 1.24E+03", sb.ToString
End Sub

Public Sub TestAppendFormatHex()
    sb.AppendFormat "{0:x}", &HFFFF
    Assert.AreEqual "ffff", sb.ToString
    
    sb.Length = 0
    sb.AppendFormat "{0:x}", CByte(255)
    Assert.AreEqual "ff", sb.ToString
End Sub

Public Sub TestAppendFormatDate()
    Dim dt As cDateTime
    Dim s As String
    
    Set dt = cDateTime.FromOADate(#1/1/2004 8:30:40 AM#)
    sb.AppendFormat "{0:dddd, dd MMMM yyyy HH':'mm':'ss}", dt
    s = sb.ToString
    Assert.AreEqual "Thursday, 01 January 2004 08:30:40", s
End Sub

Public Sub TestReturnZeroLengthString()
    Dim s As String
    
    sb.Length = 0
    s = sb.ToString
    Assert.IsFalse cString.IsNull(s), "StringBuilder should return an empty string, not a null string."
End Sub

Public Sub TestCopyTo()
    Dim ch() As Integer
    
    ReDim ch(6)
    sb.Append "hello, how are you?"
    sb.CopyTo 5, ch, 0, 7
    
    Assert.AreEqual ", how a", cString.FromCharArray(ch)
End Sub

Public Sub TestEscapeBraces()
    sb.AppendFormat "10 = {{0 }}  {0}", "hi"
    Assert.AreEqual "10 = {0 }  hi", sb.ToString
End Sub

Public Sub TestFailedCloseEscape()
    On Error GoTo errTrap
    sb.AppendFormat "10 = {0}} hi"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is FormatException, "Wrong exception caught."
End Sub

Public Sub TestFailedOpenEscape()
    On Error GoTo errTrap
    sb.AppendFormat "10 = {{0}  hi"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is FormatException, "Wrong exception caught."
End Sub

Public Sub TestMissingIndex()
    On Error GoTo errTrap
    sb.AppendFormat "10 = {}"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is FormatException, "Wrong exception caught."
End Sub


