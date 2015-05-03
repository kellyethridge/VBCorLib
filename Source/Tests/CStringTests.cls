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
            
        Case "Remove_WithValues_RemovesExpected"
            Test.Use 0, 10, ""
            Test.Use 1, 9, "1"
            Test.Use 0, 9, "0"
            Test.Use 1, 8, "10"
            Test.Use 1, 0, "1234567890"
            
        Case "Remove_WithStartIndex_ReturnsExpected"
            Test.Use 0, ""
            Test.Use 1, "1"
            Test.Use 9, "123456789"
            Test.Use 10, "1234567890"
            
        Case "Join_WihtLongsSubset_ReturnsJoined"
            Test.Use 0, 10, "1,2,3,4,5,6,7,8,9,0"
            Test.Use 1, 9, "2,3,4,5,6,7,8,9,0"
            Test.Use 0, 9, "1,2,3,4,5,6,7,8,9"
            Test.Use 1, 8, "2,3,4,5,6,7,8,9"
            
        Case "Join_WithEnumerableObject_ReturnsJoinedElements"
            Test.Use NewCollection(1, "hi", 2), "1,hi,2"
            Test.Use NewArrayList(Source:=Array(1, "hi", 2)), "1,hi,2"
            
        Case "Insert_WithValueAndIndex_ReturnsExpected"
            Test.Use 0, "ab", "ab1234567890"
            Test.Use 1, "ab", "1ab234567890"
            Test.Use 9, "ab", "123456789ab0"
            Test.Use 10, "ab", "1234567890ab"
            
        Case "IndexOfAny_WithValue_ReturnsExpected"
            Test.Use "1", 0
            Test.Use "0", 9
            Test.Use "ab5", 4
            Test.Use "abc", -1
            Test.Use NewChars("1"), 0
            Test.Use NewChars("0"), 9
            Test.Use NewChars("a", "b", "5"), 4
            Test.Use NewChars("a", "b", "c"), -1
            
        Case "IndexOfAny_WithStartIndex_ReturnsExpected"
            Test.Use "1", 1, 10
            Test.Use "0", 10, 19
            
        Case "LastIndexOfAny_WithValue_ReturnsExpected"
            Test.Use "0", 9
            Test.Use "1", 0
            
    End Select
End Sub

Public Sub IsEmptyOrWhitespace_WithValues_ReturnsExpected(ByVal s As String, ByVal Expected As Boolean)
    Dim Actual As Boolean

    Actual = CString.IsNullOrWhiteSpace(s)

    Assert.That Actual, Iz.EqualTo(Expected)
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

Public Sub ToCharArray_WithSubstringSpecified_ReturnsSubPortion()
    Dim Actual() As Integer
    
    Actual = CString.ToCharArray("hello", 1, 3)
    
    Assert.That Actual, Iz.EqualTo(NewIntegers(101, 108, 108))
End Sub

Public Sub ToCharArray_WithStartIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.ToCharArray "hello", -1, 1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub ToCharArray_WithLengthLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.ToCharArray "hello", 0, -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub ToCharArray_WithStartIndexAndLengthGreaterThanTheSizeOfCharsArray_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.ToCharArray "hello", 1, 5
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub FromCharArray_WithArrayOfCharacters_ReturnsExpected(ByRef CharactersToConvert() As Integer, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.FromCharArray(CharactersToConvert)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FromCharArray_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Chars(1 To 10) As Integer
    
    On Error Resume Next
    CString.FromCharArray Chars, 0, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub FromCharArray_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Chars(0) As Integer
    
    On Error Resume Next
    CString.FromCharArray Chars, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub FromCharArray_WithIndexAndCountBeyondUBound_ThrowsArgumentExceptin()
    Dim Chars(0 To 5) As Integer
    
    On Error Resume Next
    CString.FromCharArray Chars, 1, 6
    
    AssertArgumentException Err
End Sub

Public Sub FromCharArray_WithNullArray_ThrowsArgumentNullException()
    Dim NullChars() As Integer
    
    On Error Resume Next
    CString.FromCharArray NullChars
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub FromCharArray_WithMultiDimArray_ThrowsRankException()
    Dim MultiDim(0, 0) As Integer
    
    On Error Resume Next
    CString.FromCharArray MultiDim
    
    AssertRankException Err
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

Public Sub TrimStart_WithDefaultWhiteSpace_TrimsWhiteSpaceFromStart(ByVal s As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.TrimStart(s)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TrimStart_WithValues_ReturnsExpected(ByVal s As String, ByRef TrimChars As Variant, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.TrimStart(s, TrimChars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TrimEnd_WithDefaultWhiteSpace_TrimsWhiteSpaceFromEnd(ByVal s As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.TrimEnd(s)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TrimEnd_WithValues_ReturnsExpected(ByVal s As String, ByRef TrimChars As Variant, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.TrimEnd(s, TrimChars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Trim_WithDefaultWhiteSpace_TrimsWhiteSpace(ByVal s As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.Trim(s)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Trim_WithValues_ReturnsExpected(ByVal s As String, ByRef TrimChars As Variant, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.Trim(s, TrimChars)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Join_WhenCalled_ReturnsDelimitedString()
    Dim Value() As String
    Value = NewStrings("a", "b", "c")
    
    Dim Actual As String
    Actual = CString.Join(",", Value)
    
    Assert.That Actual, Iz.EqualTo("a,b,c")
End Sub

Public Sub Join_WithSubset_ReturnsDelimitedSubset()
    Dim Value() As String
    Value = NewStrings("a", "b", "c", "d")
    
    Dim Actual As String
    Actual = CString.Join(",", Value, 1, 2)
    
    Assert.That Actual, Iz.EqualTo("b,c")
End Sub

Public Sub Join_WithStartIndexBelowLBound_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.Join ",", Array(1, 1), -1, 2
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub Join_WithNegativeCount_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.Join ",", Array(1, 1), 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Join_WithStartIndexPlusCountIsInvalid_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.Join ",", Array(1, 2), 0, 3
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub Join_WithLongs_ReturnsLongsJoined()
    Dim Actual As String
    
    Actual = CString.Join(",", Array(1, 2, 3, 4))
    
    Assert.That Actual, Iz.EqualTo("1,2,3,4")
End Sub

Public Sub Join_WithIObjects_ReturnsJoined()
    Dim Objs(0 To 1) As Variant
    Set Objs(0) = NewInt32(32)
    Set Objs(1) = NewInt32(1)
    
    Dim Actual As String
    Actual = CString.Join(",", Objs)
    
    Assert.That Actual, Iz.EqualTo("32,1")
End Sub

Public Sub Join_WithLongArray_ReturnsJoined()
    Dim Longs() As Long
    Longs = NewLongs(1, 2, 3, 4)
    
    Dim Actual As String
    Actual = CString.Join(",", Longs)
    
    Assert.That Actual, Iz.EqualTo("1,2,3,4")
End Sub

Public Sub Join_WihtLongsSubset_ReturnsJoined(ByVal StartIndex As Long, ByVal Count As Long, ByVal Expected As String)
    Dim Longs() As Long
    Longs = NewLongs(1, 2, 3, 4, 5, 6, 7, 8, 9, 0)
    
    Dim Actual As String
    Actual = CString.Join(",", Longs, StartIndex, Count)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Join_WithEnumerableObject_ReturnsJoinedElements(ByVal Value As Variant, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.Join(",", Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Join_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    
    CString.Join ",", Nothing
    
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub Join_WithNonEnumerableObject_ThrowsArgumentException()
    On Error Resume Next
    
    CString.Join ",", Me
    
    AssertArgumentException Err, "Value"
End Sub

Public Sub Remove_WithValues_RemovesExpected(ByVal Index As Long, ByVal Count As Long, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.Remove("1234567890", Index, Count)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Remove_WithStartIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.Remove "a", -1, 1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub Remove_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.Remove "a", 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub Remove_WithStartIndexAndCountInvalidPosition_ThrowsArgumentException()
    On Error Resume Next
    
    CString.Remove "a", 0, 2
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub Remove_WithStartIndex_ReturnsExpected(ByVal StartIndex As Long, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.Remove("1234567890", StartIndex)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Insert_WithValueAndIndex_ReturnsExpected(ByVal StartIndex As Long, ByVal Value As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CString.Insert("1234567890", StartIndex, Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Insert_WithStartIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.Insert "ab", -1, "a"
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub Insert_WithStartIndexAfterLengthOfString_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.Insert "ab", 3, "a"
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub IndexOfAny_WithValue_ReturnsExpected(ByRef OfAny As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = CString.IndexOfAny("1234567890", OfAny)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IndexOfAny_WithStartIndex_ReturnsExpected(ByRef OfAny As Variant, ByVal StartIndex As Long, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = CString.IndexOfAny("12345678901234567890", OfAny, StartIndex)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IndexOfAny_WithStartIndexAndCount_ReturnsExpected()
    Dim Actual As Long
    
    Actual = CString.IndexOfAny("12345678901234567890", "5", 0, 10)
    
    Assert.That Actual, Iz.EqualTo(4)
End Sub

Public Sub IndexOfAny_WithStartIndexMissingAndHasCount_ThrowsArgumentException()
    On Error Resume Next
    
    CString.IndexOfAny "a", "a", Count:=1
    
    AssertArgumentException Err, "StartIndex"
End Sub

Public Sub IndexOfAny_WithStartIndexLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.IndexOfAny "a", "a", -1
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub IndexOfAny_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.IndexOfAny "a", "a", 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub IndexOfAny_WithStartIndexGreaterThanTheLengthOfStringToSearch_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.IndexOfAny "a", "a", 2
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub IndexOfAny_WithStartIndexPlusCountGreaterThanLengthOfString_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CString.IndexOfAny "a", "a", 0, 2
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub LastIndexOfAny_WithValue_ReturnsExpected(ByRef OfAny As Variant, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = CString.LastIndexOfAny("1234567890", OfAny)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub




Private Property Get WhiteSpaces() As String
    WhiteSpaces = ChrW$(&H20) & ChrW$(&HD) & ChrW$(&H9) & ChrW$(&HA) & ChrW$(&HB) & ChrW$(&HC) & ChrW$(&H85) & ChrW$(&HA0) & ChrW$(&H1680) & ChrW$(&H180E) & _
        ChrW$(&H2000) & ChrW$(&H2001) & ChrW$(&H2002) & ChrW$(&H2003) & ChrW$(&H2004) & ChrW$(&H2005) & ChrW$(&H2006) & ChrW$(&H2007) & ChrW$(&H2008) & ChrW$(&H2009) & ChrW$(&H200A) & _
        ChrW$(&H2028) & ChrW$(&H2029) & ChrW$(&H202F) & ChrW$(&H205F) & _
        ChrW$(&H3000)
End Property

Private Property Get WhiteSpaceChars() As Integer()
    WhiteSpaceChars = NewIntegers(&H20, &HD, &H9, &HA, &HB, &HC, &H85, &HA0, &H1680, &H180E, _
        &H2000, &H2001, &H2002, &H2003, &H2004, &H2005, &H2006, &H2007, &H2008, &H2009, &H200A, _
        &H2028, &H2029, &H202F, &H205F, _
        &H3000)
End Property



