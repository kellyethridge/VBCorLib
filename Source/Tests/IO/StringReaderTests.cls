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

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
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

'Public Sub TestReadBlock()
'    Dim ch(100) As Integer
'    Dim i As Long
'
'    Set sr = NewStringReader("abcdefghijk")
'
'    Assert.AreEqual 11, sr.ReadBlock(ch, 0, 100)
'    For i = 0 To 10
'        Assert.AreEqual Asc("a") + i, ch(i)
'    Next i
'End Sub
'
'Public Sub TestReadBlockPastEndOfStream()
'    Dim ch(10) As Integer
'    Dim i As Long
'
'    Set sr = Cor.NewStringReader("12345678")
'    Assert.AreEqual 8, sr.ReadBlock(ch, 0, 11)
'End Sub
'
'Public Sub TestReadLine()
'    Set sr = NewStringReader("abc" & vbCr & vbLf & "def" & vbLf & vbCr)
'    Assert.AreEqual "abc", sr.ReadLine
'    Assert.AreEqual "def", sr.ReadLine
'    Assert.IsFalse CorString.IsNull(sr.ReadLine)
'    Assert.IsTrue CorString.IsNull(sr.ReadLine)
'End Sub
'
'Public Sub TestReadToEnd()
'    Set sr = NewStringReader("abcdefghi")
'    sr.Read
'    Assert.AreEqual "bcdefghi", sr.ReadToEnd
'End Sub
