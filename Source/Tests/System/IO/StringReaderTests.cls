VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringReaderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ReadLine_WithValues_ReturnsExpected"
            Test.Use "abcdef", "abcdef"
            Test.Use "abcdef" & vbCr, "abcdef"
            Test.Use "abcdef" & vbLf, "abcdef"
            Test.Use "abcdef" & vbCrLf, "abcdef"
            Test.Use vbCr & "abcdef", ""
            Test.Use vbLf & "abcdef", ""
            Test.Use vbCrLf & "abcdef", ""
            Test.Use "abc" & vbCr & "def", "abc"
            Test.Use "abc" & vbLf & "def", "abc"
            Test.Use "abc" & vbCrLf & "def", "abc"
            
    End Select
End Sub

Public Sub Read_WithEmptyString_ReturnsMinusOne()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("")
    
    Dim Actual As Long
    Actual = Reader.Read
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub Read_WhenCalled_ReturnsExpected()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abc")
    
    Assert.That Reader.Read, Iz.EqualTo(Asc("a"))
    Assert.That Reader.Read, Iz.EqualTo(Asc("b"))
    Assert.That Reader.Read, Iz.EqualTo(Asc("c"))
    Assert.That Reader.Read, Iz.EqualTo(-1)
End Sub

Public Sub Peek_WithEmptyString_ReturnsMinusOne()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("")
    
    Dim Actual As Long
    Actual = Reader.Peek
    
    Assert.That Actual, Iz.EqualTo(-1)
End Sub

Public Sub Peek_WhenCalled_ReturnsValue()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("ab")
    
    Dim Actual As Long
    Actual = Reader.Peek
    
    Assert.That Actual, Iz.EqualTo(Asc("a"))
End Sub

Public Sub Peek_WhenCalled_DoesNotAdvancePosition()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("ab")
    
    Reader.Peek
    
    Dim Actual As Long
    Actual = Reader.Read
    Assert.That Actual, Iz.EqualTo(Asc("a"))
End Sub

Public Sub ReadBlock_WithNullChars_ThrowsArgumentNullException()
    Dim NullChars() As Integer
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abc")
    
    On Error Resume Next
    Reader.ReadBlock NullChars, 0, 1
    
    AssertArgumentNullException Err, "Buffer"
End Sub

Public Sub ReadBlock_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Chars(2 To 5) As Integer
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abc")
    
    On Error Resume Next
    Reader.ReadBlock Chars, 1, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub ReadBlock_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Chars(0 To 1) As Integer
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abc")
    
    On Error Resume Next
    Reader.ReadBlock Chars, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub ReadBlock_WithIndexPlusCountInvalid_ThrowsArgumentException()
    Dim Chars(0 To 1) As Integer
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abc")
    
    On Error Resume Next
    Reader.ReadBlock Chars, 1, 2
    
    AssertArgumentException Err
End Sub

Public Sub ReadBlock_WhenCalled_FillsBuffer()
    Dim Chars(0 To 1) As Integer
    Dim Count As Long
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abc")
    
    Count = Reader.ReadBlock(Chars, 0, 2)
    
    Assert.That Count, Iz.EqualTo(2)
    Assert.That Chars, Iz.EqualTo(NewChars("a", "b"))
End Sub

Public Sub ReadBlock_WithRequestLargerThanRemainingInReader_ReturnsRemaining()
    Dim Chars(0 To 5) As Integer
    Dim Count As Long
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abcd")
    
    Count = Reader.ReadBlock(Chars, 0, 6)
    
    Assert.That Count, Iz.EqualTo(4)
    Assert.That Chars, Iz.EqualTo(NewChars("a", "b", "c", "d", vbNullChar, vbNullChar))
End Sub

Public Sub ReadBlock_WhenPositionChanges_ReadsFromCurrentPosition()
    Dim Chars(0 To 1) As Integer
    Dim Count As Long
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abc")
    Reader.Read
    
    Count = Reader.ReadBlock(Chars, 0, 2)
    
    Assert.That Count, Iz.EqualTo(2)
    Assert.That Chars, Iz.EqualTo(NewChars("b", "c"))
End Sub

Public Sub ReadLine_WithValues_ReturnsExpected(ByVal s As String, ByVal Expected As String)
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader(s)
    
    Dim Actual As String
    Actual = Reader.ReadLine
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadLine_WhenAllTextHasBeenRead_ReturnsNullString()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abc")
    Reader.ReadLine
    
    Dim Actual As String
    Actual = Reader.ReadLine
    
    Assert.That CorString.IsNull(Actual), Iz.True
End Sub

Public Sub ReadLine_WithMultipleBreaks_ReturnsIndividualLines()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("a" & vbCr & "b" & vbCrLf & "c" & vbLf & "d")
    
    Assert.That Reader.ReadLine, Iz.EqualTo("a")
    Assert.That Reader.ReadLine, Iz.EqualTo("b")
    Assert.That Reader.ReadLine, Iz.EqualTo("c")
    Assert.That Reader.ReadLine, Iz.EqualTo("d")
    Assert.That CorString.IsNull(Reader.ReadLine), Iz.True
End Sub

Public Sub ReadLine_WhenPositionChanges_ReadsFromCurrentPosition()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abcd" & vbCr & "ef")
    Reader.Read
    
    Dim Actual As String
    Actual = Reader.ReadLine
    
    Assert.That Actual, Iz.EqualTo("bcd")
End Sub

Public Sub ReadLine_WithEmptyString_ReturnsNullString()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("")
    
    Dim Actual As String
    Actual = Reader.ReadLine
    
    Assert.That CorString.IsNull(Actual), Iz.True
End Sub

Public Sub ReadToEnd_WhenCalled_ReturnsRemainingCharacters()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abcde")
    
    Dim Actual As String
    Actual = Reader.ReadToEnd
    
    Assert.That Actual, Iz.EqualTo("abcde")
End Sub

Public Sub ReadToEnd_WithReadToEndAlreadyCalled_ReturnsNullString()
    Dim Reader As StringReader
    Set Reader = Cor.NewStringReader("abcd")
    Reader.ReadToEnd
    
    Dim Actual As String
    Actual = Reader.ReadToEnd
    
    Assert.That CorString.IsNull(Actual), Iz.True
End Sub

