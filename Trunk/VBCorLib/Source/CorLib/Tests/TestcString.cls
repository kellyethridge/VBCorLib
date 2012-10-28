VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestcString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestFormat()
    Assert.AreEqual "hel1234lo", cString.Format("hel{0}lo", 1234)
End Sub

Public Sub TestToCharArray()
    Dim s As String
    Dim ch() As Integer
    
    s = "Hello"
    ch = cString.ToCharArray(s)
    Assert.AreEqual 5, cArray.GetLength(ch, 1), "Array length should be 5."
    Assert.AreEqual Asc("H"), ch(0)
    Assert.AreEqual Asc("e"), ch(1)
    Assert.AreEqual Asc("l"), ch(2)
    Assert.AreEqual Asc("l"), ch(3)
    Assert.AreEqual Asc("o"), ch(4)
End Sub

Public Sub TestFromCharArray()
    Dim Chars() As Integer
    ReDim Chars(25)
    Dim i As Long
    For i = 0 To 25
        Chars(i) = i + 65
    Next i
    
    Dim s As String
    s = cString.FromCharArray(Chars)
    Assert.AreEqual "ABCDEFGHIJKLMNOPQRSTUVWXYZ", s
End Sub

Public Sub TestFromEmptyCharArray()
    Dim Chars() As Integer
    Chars = Cor.NewIntegers()
    
    Assert.AreEqual "", cString.FromCharArray(Chars)
End Sub

Public Sub TestEmptyToCharArray()
    Dim s As String
    Dim ch() As Integer
    
    ch = cString.ToCharArray(s)
    Assert.AreEqual 0, cArray.GetLength(ch), "Array should be empty."
End Sub

Public Sub TestPadLeft()
    Assert.AreEqual "   abc", cString.PadLeft("abc", 6)
End Sub

Public Sub TestPadLeftWithCharCode()
    Assert.AreEqual "000abc", cString.PadLeft("abc", 6, 48)
End Sub

Public Sub TestPadLeftTooShort()
    Assert.AreEqual "abc", cString.PadLeft("abc", 2)
End Sub

Public Sub TestPadRight()
    Assert.AreEqual "abc   ", cString.PadRight("abc", 6)
End Sub

Public Sub TestPadRightWithCharCode()
    Assert.AreEqual "abc000", cString.PadRight("abc", 6, 48)
End Sub

Public Sub TestPadRightTooShort()
    Assert.AreEqual "abc", cString.PadRight("abc", 2)
End Sub

Public Sub TestTrimStart()
    Dim Chars() As Integer
    Dim s As String
    
    Chars = cString.ToCharArray("abc")
    s = "cabhi"
    s = cString.TrimStart(s, Chars)
    Assert.AreEqual "hi", s
End Sub

Public Sub TestTrimStartNoTrim()
    Dim Chars(0) As Integer
    Dim s As String
    
    Chars(0) = 48
    s = "hello"
    s = cString.TrimStart(s, Chars)
    Assert.AreEqual "hello", s
End Sub

Public Sub TestTrimStartAllTrimmed()
    Dim Chars(0) As Integer
    Dim s As String
    
    Chars(0) = 48
    s = "00000"
    s = cString.TrimStart(s, Chars)
    Assert.AreEqual "", s
End Sub

Public Sub TestTrimEnd()
    Dim Chars() As Integer
    Dim s As String
    
    Chars = cString.ToCharArray("abc")
    s = "hicab"
    s = cString.TrimEnd(s, Chars)
    Assert.AreEqual "hi", s
End Sub

Public Sub TestTrimEndNoTrim()
    Dim Chars(0) As Integer
    Dim s As String
    
    Chars(0) = 48
    s = "hello"
    s = cString.TrimEnd(s, Chars)
    Assert.AreEqual "hello", s
End Sub

Public Sub TestTrimEndAllTrimmed()
    Dim Chars(0) As Integer
    Dim s As String
    
    Chars(0) = 48
    s = "00000"
    s = cString.TrimEnd(s, Chars)
    Assert.AreEqual "", s
End Sub

Public Sub TestTrim()
    Dim s As String
    
    s = "*;|@123***456@|;*"
    s = cString.Trim(s, "*;|@")
    Assert.AreEqual "123***456", s
End Sub

Public Sub TestTrimWhiteSpace()
    Dim s As String
    
    s = vbCr & vbCr & " hi " & vbLf & vbLf
    s = cString.Trim(s)
    Assert.AreEqual "hi", s
End Sub

Public Sub TestTrimSpaces()
    Dim s As String
    
    s = "     "
    s = cString.Trim(s)
    Assert.AreEqual "", s
End Sub

Public Sub TestTrimEmptyString()
    Dim s As String
    s = cString.Trim(s)
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
    
    Assert.AreEqual "bye,again", cString.Join(v, ",", 1)
End Sub

Public Sub TestRemove()
    Dim s As String
    
    s = "1234567890"
    Assert.AreEqual "1234890", cString.Remove(s, 4, 3)
End Sub

Public Sub TestRemoveBeginning()
    Dim s As String
    
    s = "1234567890"
    Assert.AreEqual "890", cString.Remove(s, 0, 7)
End Sub

Public Sub TestRemoveEnd()
    Dim s As String
    
    s = "1234567890"
    Assert.AreEqual "123", cString.Remove(s, 3, 7)
End Sub

Public Sub TestIndexOfAnyString()
    Assert.AreEqual 1, cString.IndexOfAny("abc", "b")
End Sub

Public Sub TestIndexOfAnyChars()
    Assert.AreEqual 0, cString.IndexOfAny("abc", NewArray(ciInteger, 98, 97))
End Sub

Public Sub TestIndexOfAnyNotFound()
    Assert.AreEqual -1, cString.IndexOfAny("abc", "123")
End Sub

Public Sub TestIndexOfAnySubString()
    Assert.AreEqual 1, cString.IndexOfAny("abc", "b", 1, 1)
End Sub

Public Sub TestLastIndexOfAnyString()
    Assert.AreEqual 2, cString.LastIndexOfAny("abc", "bc")
End Sub

Public Sub TestLastIndexOfAnyNone()
    Assert.AreEqual -1, cString.LastIndexOfAny("abc", "123")
End Sub

Public Sub TestInsertInMiddle()
    Assert.AreEqual "abc123def", cString.Insert("abcdef", 3, "123")
End Sub

Public Sub TestInsertAtBeginning()
    Assert.AreEqual "123abcdef", cString.Insert("abcdef", 0, "123")
End Sub

Public Sub TestInsertAtEnd()
    Assert.AreEqual "abcdef123", cString.Insert("abcdef", 6, "123")
End Sub

Public Sub TestInsertInvalidIndexBefore()
    On Error GoTo errTrap
    Call cString.Insert("abc", -1, "123")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentOutOfRangeException, "Wrong exception type caught."
End Sub

Public Sub TestInsertInvalidIndexAfter()
    On Error GoTo errTrap
    Call cString.Insert("abc", 4, "123")
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
    
    Assert.AreEqual expectedvalue, cString.Join(s, Delimiter, Index, Count)
End Sub

