VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF7GetChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF7GetChars
'
Option Explicit
Implements ICategorizable

Private En As UTF7Encoding
Private b() As Byte
Private ch() As Integer

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub Setup()
    Set En = New UTF7Encoding
End Sub

Public Sub TestPlus()
    ReDim ch(10)
    
    b = NewBytes(43)
    Assert.AreEqual 0, En.GetCharsEx(b, 0, 1, ch, 0)
    Assert.AreEqual 0, ch(0)
End Sub

Public Sub TestOneNonDirect()
    ReDim ch(100)
    b = NewBytes(33)
    Assert.AreEqual 1, En.GetCharsEx(b, 0, 1, ch, 0)
    Assert.AreEqual 33, ch(0)
End Sub

Public Sub TestPlusMinus()
    ReDim ch(100)
    b = NewBytes(43, 45)
    Assert.AreEqual 1, En.GetCharsEx(b, 0, 2, ch, 0)
    Assert.AreEqual 43, ch(0)
    Assert.AreEqual 0, ch(1)
End Sub

Public Sub TestSimple()
    ReDim ch(100)
    b = En.GetBytes("hello")
    Assert.AreEqual 5, En.GetCharsEx(b, 0, UBound(b) + 1, ch, 0)
    Assert.AreEqual "hello", CorString.FromCharArray(ch, 0, 5)
    Assert.AreEqual 0, ch(5)
End Sub

Public Sub TestJapaneseWord_nihongo()
    Dim ch() As Integer
    ch = NewIntegers(&H65E5, &H672C, &H8A9E)
    b = En.GetBytes(ch)
    ReDim ch(100)
    Assert.AreEqual 3, En.GetCharsEx(b, 0, UBound(b) + 1, ch, 0)
    Assert.AreEqual &H65E5, ch(0)
    Assert.AreEqual &H672C, ch(1)
    Assert.AreEqual &H8A9E, ch(2)
    Assert.AreEqual 0, ch(3)
End Sub

Public Sub TestGetString()
    b = En.GetBytes("Hello there.")
    Assert.AreEqual "Hello there.", En.GetString(b)
End Sub

Public Sub TestDecoder()
    Dim c As Decoder
    Dim i As Long
    
    b = En.GetBytes("a" & ChrW$(&H65E5) & "a")
    ReDim ch(10)
    
    Set c = En.GetDecoder
    For i = 0 To 6
        Assert.AreEqual Choose(i + 1, 1, 0, 0, 0, 1, 0, 1), c.GetChars(b, i, 1, ch, 0), "Failed on index " & i & "."
    Next i
    
    Set c = En.GetDecoder
    For i = 0 To 6
        Assert.AreEqual Choose(i + 1, 1, 1, 1, 1, 2, 2, 3), c.GetChars(b, 0, i + 1, ch, 0), "Failed on index " & i & "."
    Next i
End Sub

Public Sub TestDecoderWith1001CharsFromFile()
    Dim FF As Long
    Dim b() As Byte
    Dim ch() As Integer
    Dim i As Long
    Dim En As UTF7Encoding
    Dim c As Decoder
    Dim Char As Long
    
    FF = FreeFile
    Open App.Path & "\UTF7_Chars_0to1000.txt" For Binary As #FF
    ReDim b(0 To LOF(FF) - 1)
    Get #FF, , b
    Close #FF
    
    Set En = New UTF7Encoding
    Set c = En.GetDecoder
    
    ReDim ch(0)
    For i = 0 To UBound(b)
        If c.GetChars(b, i, 1, ch, 0) = 1 Then
            Assert.AreEqual Char, ch(0)
            Char = Char + 1
        End If
    Next i
End Sub

Public Sub TestGet1001CharsFromfile()
    Dim FF As Long
    Dim b() As Byte
    Dim ch() As Integer
    Dim i As Long
    Dim En As UTF7Encoding
    
    FF = FreeFile
    Open App.Path & "\UTF7_Chars_0to1000.txt" For Binary As #FF
    ReDim b(0 To LOF(FF) - 1)
    Get #FF, , b
    Close #FF
    
    Set En = New UTF7Encoding
    
    Assert.AreEqual 1001, En.GetCharCount(b), "GetCharCount should return 1001."
    
    ch = En.GetChars(b)
    Assert.AreEqual 1001, CorArray.Length(ch), "Length should be 1001."
    
    For i = 0 To 1000
        Assert.AreEqual i, ch(i), "Wrong char at index " & i & "."
    Next i
    
End Sub

Public Sub TestEmptyArray()
    Dim b() As Byte
    Dim ch() As Integer
    
    b = CorArray.CreateInstance(vbByte)
    ch = En.GetChars(b)
    Assert.IsFalse CorArray.IsNull(ch)
    Assert.AreEqual 0, CorArray.Length(ch)
    
    ReDim ch(10)
    Assert.AreEqual 0, En.GetCharsEx(b, 0, 0, ch, 0)
End Sub
