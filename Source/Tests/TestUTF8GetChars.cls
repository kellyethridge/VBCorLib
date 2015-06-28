VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF8GetChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF8GetChars
'
Option Explicit
Implements ICategorizable

Private e As UTF8Encoding
Private b() As Byte
Private ch() As Integer

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub Setup()
    Set e = NewUTF8Encoding(False)
End Sub

Public Sub Teardown()
    Erase b
    Erase ch
End Sub

Public Sub TestFrom1ByteAnsii()
    Dim ch() As Integer
    ch = e.GetChars(NewBytes(65))
    Assert.AreEqual 1, CorArray.Length(ch)
    Assert.AreEqual 65, ch(0)
End Sub

Public Sub TestFrom10BytesAnsii()
    Dim ch() As Integer
    Dim i As Long
    
    b = e.GetBytes(String$(10, &H7F))
    ch = e.GetChars(b)
    Assert.AreEqual 10, CorArray.Length(ch)
    For i = 0 To 9
        Assert.AreEqual &H7F, ch(i)
    Next i
End Sub

Public Sub TestSurrogatePair()
    b = e.GetBytes(NewIntegers(&HDBFF, &HDC00))
    ch = e.GetChars(b)
    Assert.AreEqual 2, CorArray.Length(ch)
    Assert.AreEqual &HDBFF, ch(0)
    Assert.AreEqual &HDC00, ch(1)
End Sub

Public Sub TestNonAscii()
    b = e.GetBytes(String$(10, 167))
    ch = e.GetChars(b)
    Assert.AreEqual 10, CorArray.Length(ch)
    
    Assert.AreEqual String$(10, 167), CorString.FromCharArray(ch)
End Sub

Public Sub TestComboChars()
    Dim asciiChars As String
    Dim nonAsciiChars As String
    Dim surrogateChars As String
    Dim i As Long
    
    asciiChars = String$(10, "A")
    nonAsciiChars = String$(10, ChrW$(167))
    For i = 1 To 10
        surrogateChars = surrogateChars & ChrW$(&HD800) & ChrW$(&HDC00)
    Next i
    Dim s As String
    s = nonAsciiChars & surrogateChars & asciiChars
    
    b = e.GetBytes(s)
    ch = e.GetChars(b)
    Assert.AreEqual 40, CorArray.Length(ch)
    Assert.AreEqual s, CorString.FromCharArray(ch)
End Sub

Public Sub TestGetStringComboChars()
    Dim asciiChars As String
    Dim nonAsciiChars As String
    Dim surrogateChars As String
    Dim i As Long
    
    asciiChars = String$(10, "A")
    nonAsciiChars = String$(10, ChrW$(167))
    For i = 1 To 10
        surrogateChars = surrogateChars & ChrW$(&HD801) & ChrW$(&HDC01)
    Next i
    Dim s As String
    s = nonAsciiChars & surrogateChars & asciiChars
    
    b = e.GetBytes(s)
    Dim s2 As String
    s2 = e.GetString(b)
    Assert.AreEqual 40, Len(s)
    Assert.AreEqual s, s
End Sub

Public Sub TestCharBufferTooSmall()
    Dim ch() As Integer
    b = e.GetBytes("Hello")
    ReDim ch(3)
    
    On Error GoTo errTrap
    e.GetCharsEx b, 0, 5, ch, 0
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Incorrect exception type."
End Sub

Public Sub TestShortestFrom()
    Dim ch() As Integer
    
    b = e.GetBytes(ChrW$(8191))
    Assert.AreEqual 3, CorArray.Length(b)
    ch = e.GetChars(b)
    Assert.AreEqual 1, CorArray.Length(ch)
    Assert.AreEqual 8191, ch(0)
End Sub

Public Sub TestLongBuffer()
    b = e.GetBytes("hello")
    ReDim Preserve b(0 To 100)
    Assert.AreEqual "hello", e.GetString(b, 0, 5)
End Sub

Public Sub TestGet1001CharsFromfile()
    Dim FF As Long
    Dim b() As Byte
    Dim ch() As Integer
    Dim i As Long
    Dim En As UTF8Encoding
    
    FF = FreeFile
    Open App.Path & "\UTF8_Chars_0to1000.txt" For Binary As #FF
    ReDim b(0 To LOF(FF) - 1)
    Get #FF, , b
    Close #FF
    
    Set En = New UTF8Encoding
    ch = En.GetChars(b)
    Assert.AreEqual 1001, CorArray.Length(ch)
    
    For i = 0 To 1000
        Assert.AreEqual i, ch(i)
    Next i
End Sub

Public Sub TestBadSurrogatePair()
    Dim ch() As Integer
    
    b = e.GetBytes(ChrW$(&HD800) & ChrW$(&HD801))
    ch = e.GetChars(b)
    Assert.AreEqual &HD800, ch(0)
    Assert.AreEqual &HD801, ch(1)
End Sub

Public Sub TestZeroLengthArray()
    Dim ch() As Integer
    
    b = CorArray.CreateInstance(vbByte)
    ch = e.GetChars(b)
    Assert.IsFalse CorArray.IsNull(ch), "Array should not be null."
    Assert.AreEqual 0, CorArray.Length(ch)
    
End Sub
