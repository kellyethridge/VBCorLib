VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF8Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF8Encoding
'
Option Explicit
Implements ICategorizable

Private e As UTF8Encoding
Private s As String
Private b() As Byte

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub Setup()
    Set e = NewUTF8Encoding(True)
End Sub

Public Sub TestGetPreamble()
    Dim b() As Byte
    
    b = e.GetPreamble
    Assert.IsFalse CorArray.IsNull(b)
    Assert.AreEqual 3, CorArray.Length(b)
    '0xEF, 0xBB, 0xBF
    Assert.AreEqual &HEF, b(0)
    Assert.AreEqual &HBB, b(1)
    Assert.AreEqual &HBF, b(2)
End Sub

Public Sub TestGetBytesFrom1CharString()
    s = Chr$(&H7F)
    b = e.GetBytes(s)
    Assert.AreEqual 1, CorArray.Length(b)
    Assert.AreEqual &H7F, b(0)
End Sub

Public Sub TestGetBytesFrom100CharString()
    s = String$(100, &H7F)
    b = e.GetBytes(s)
    Assert.AreEqual 100, CorArray.Length(b)
    
    Dim i As Long
    For i = 0 To 99
        Assert.AreEqual &H7F, b(i)
    Next i
End Sub

Public Sub TestGetBytesFromSurrogatePair()
    Dim ch() As Integer
    
    ch = NewIntegers(&HD800, &HDC00)
    b = e.GetBytes(ch)
    Assert.AreEqual 4, CorArray.Length(b)
    Assert.AreEqual 240, b(0)
    Assert.AreEqual 144, b(1)
    Assert.AreEqual 128, b(2)
    Assert.AreEqual 128, b(3)
End Sub

Public Sub TestGetBytesFromHighSurrogateOnly()
    Dim ch() As Integer
    
    ch = NewIntegers(&HD801)
    b = e.GetBytes(ch)
    Assert.AreEqual 3, CorArray.Length(b)
    Assert.AreEqual 237, b(0)
    Assert.AreEqual 160, b(1)
    Assert.AreEqual 129, b(2)
End Sub

Public Sub TestGetBytesFromLowSurrogateOnly()
    b = e.GetBytes(ChrW$(&HDC01))
    Assert.AreEqual 3, CorArray.Length(b)
    Assert.AreEqual 237, b(0)
    Assert.AreEqual 176, b(1)
    Assert.AreEqual 129, b(2)
End Sub

Public Sub TestGetBytesExFromHighSurrogateOnly()
    Dim ch() As Integer
    
    ch = NewIntegers(&HD801)
    ReDim b(0 To e.GetByteCount(ch) - 1)
    Assert.AreEqual 3, e.GetBytesEx(ch, 0, 1, b, 0)
    Assert.AreEqual 3, CorArray.Length(b)
    Assert.AreEqual 237, b(0)
    Assert.AreEqual 160, b(1)
    Assert.AreEqual 129, b(2)
End Sub

Public Sub TestGetBytesExFromLowSurrogateOnly()
    ReDim b(0 To e.GetByteCount(ChrW$(&HDC01)) - 1)
    Assert.AreEqual 3, e.GetBytesEx(ChrW$(&HDC01), 0, 1, b, 0)
    Assert.AreEqual 3, CorArray.Length(b)
    Assert.AreEqual 237, b(0)
    Assert.AreEqual 176, b(1)
    Assert.AreEqual 129, b(2)
End Sub

Public Sub TestGetBytesFromAll()
    b = e.GetBytes(NewIntegers(97, &HD801, &HDC01, &H600))
    Assert.AreEqual 7, CorArray.Length(b)
    Assert.AreEqual 97, b(0)
    Assert.AreEqual 240, b(1)
    Assert.AreEqual 144, b(2)
    Assert.AreEqual 144, b(3)
    Assert.AreEqual 129, b(4)
    Assert.AreEqual 216, b(5)
    Assert.AreEqual 128, b(6)
End Sub

Public Sub TestGetBytesFromSubArray()
    b = e.GetBytes(NewIntegers(97, &HD801, &HDC01, &H600), 1, 2)
    Assert.AreEqual 4, CorArray.Length(b)
    Assert.AreEqual 240, b(0)
    Assert.AreEqual 144, b(1)
    Assert.AreEqual 144, b(2)
    Assert.AreEqual 129, b(3)
End Sub

Public Sub TestGetBytesEx()
    ReDim b(6)
    Assert.AreEqual 7, e.GetBytesEx(NewIntegers(97, &HD801, &HDC01, &H600), 0, 4, b, 0)
    Assert.AreEqual 97, b(0)
    Assert.AreEqual 240, b(1)
    Assert.AreEqual 144, b(2)
    Assert.AreEqual 144, b(3)
    Assert.AreEqual 129, b(4)
    Assert.AreEqual 216, b(5)
    Assert.AreEqual 128, b(6)
End Sub

Public Sub TestGetBytesExBytesTooSmall()
    ReDim b(5)
    
    On Error GoTo errTrap
    e.GetBytesEx NewIntegers(97, &HD801, &HDC01, &H600), 0, 4, b, 0
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
End Sub

Public Sub TestEncoder()
    Dim En As Encoder
    ReDim b(3)
    Dim ch(0) As Integer
    
    Set En = e.GetEncoder
    
    ch(0) = &HD801
    Assert.AreEqual 0, En.GetBytes(ch, 0, 1, b, 0, False)
    ch(0) = &HDC02
    Assert.AreEqual 4, En.GetBytes(ch, 0, 1, b, 0, False)
    Assert.AreEqual 240, b(0)
    Assert.AreEqual 144, b(1)
    Assert.AreEqual 144, b(2)
    Assert.AreEqual 130, b(3)
End Sub

Public Sub TestDecoder()
    Dim de As Decoder
    Dim ch(1) As Integer
    
    Set de = e.GetDecoder
    b = e.GetBytes(ChrW$(&HD801) & ChrW$(&HDC02))
    Assert.AreEqual 0, de.GetChars(b, 0, 1, ch, 0)
    Assert.AreEqual 2, de.GetChars(b, 1, 3, ch, 0)
    Assert.AreEqual &HD801, ch(0)
    Assert.AreEqual &HDC02, ch(1)
End Sub

Public Sub TestEncodeChars0To1000()
    Dim ch() As Integer
    Dim i As Long
    Dim b() As Byte
    Dim fileBytes() As Byte
    Dim FF As Integer
    
    ReDim ch(1000)
    For i = 0 To 1000
        ch(i) = i
    Next i
    b = e.GetBytes(ch)
    
    Assert.AreEqual 1874, CorArray.Length(b)
    FF = FreeFile
    Open App.Path & "\UTF8_Chars_0to1000.txt" For Binary As #FF
    ReDim fileBytes(0 To LOF(FF) - 1)
    Get #FF, , fileBytes
    Close #FF
    
    For i = 0 To CorArray.Length(b) - 1
        Assert.AreEqual fileBytes(i), b(i)
    Next i
End Sub

Public Sub TestGetBytesFromHugeString()
    Dim s As String
    Dim b() As Byte
    Dim i As Long
    
    s = String$(100000, "1")
    Assert.AreEqual 100000, e.GetByteCount(s)
    b = e.GetBytes(s)
    Assert.AreEqual 100000, CorArray.Length(b)
    For i = 0 To UBound(b)
        Assert.AreEqual 49, b(i)
    Next i
End Sub

Public Sub TestEncodingEmptyString()
    Dim b() As Byte
    
    b = e.GetBytes("")
    Assert.IsFalse CorArray.IsNull(b), "Array should not be null."
    Assert.AreEqual 0, CorArray.Length(b)

    ReDim b(10)
    Assert.AreEqual 0, e.GetBytesEx("", 0, 0, b, 0)
End Sub


