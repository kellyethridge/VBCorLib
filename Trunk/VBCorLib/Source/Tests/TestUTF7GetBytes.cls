VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF7GetBytes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF7GetBytes
'
Option Explicit
Implements ICategorizable

Private e As UTF7Encoding
Private b() As Byte
Private ch() As Integer

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub Setup()
    Set e = New UTF7Encoding
End Sub

Public Sub TestDirectlyEncodable()
    ch = NewIntegers(65)
    b = e.GetBytes(ch)
    Assert.AreEqual 1, CorArray.GetLength(b)
    Assert.AreEqual 65, b(0)
End Sub

Public Sub TestEncodeSinglePlus()
    ch = NewIntegers(43)
    b = e.GetBytes(ch)
    Assert.AreEqual 2, CorArray.GetLength(b)
    Assert.AreEqual 43, b(0)
    Assert.AreEqual 45, b(1)
End Sub

Public Sub TestOneNonDirectlyEncodable()
    ch = NewIntegers(33)
    b = e.GetBytes(ch)
    Assert.AreEqual 5, CorArray.GetLength(b)
    Assert.AreEqual 43, b(0)
    Assert.AreEqual 65, b(1)
    Assert.AreEqual 67, b(2)
    Assert.AreEqual 69, b(3)
    Assert.AreEqual 45, b(4)
End Sub

Public Sub TestTwoNonDirectlyEncodable()
    ch = NewIntegers(33, 33)
    b = e.GetBytes(ch)
    
    Assert.AreEqual 8, CorArray.GetLength(b)
    
    Dim i As Long
    For i = 0 To 7
        Assert.AreEqual Choose(i + 1, 43, 65, 67, 69, 65, 73, 81, 45), b(i)
    Next i
End Sub

Public Sub TestSimpleCombo()
    ch = NewIntegers(&H41, &H2262, &H391, &H2E)
    b = e.GetBytes(ch)
    
    Assert.AreEqual 10, CorArray.GetLength(b)
    Dim i As Long
    For i = 0 To 9
        Assert.AreEqual Choose(i + 1, 65, 43, 73, 109, 73, 68, 107, 81, 45, 46), b(i)
    Next i
End Sub

Public Sub TestSimpleCombo2()
    ch = NewIntegers(&H48, &H69, &H20, &H4D, &H6F, &H6D, &H20, &H2D, &H263A, &H2D, &H21)
    b = e.GetBytes(ch)
    
    Assert.AreEqual 19, CorArray.GetLength(b)
    Dim i As Long
    For i = 0 To 18
        Assert.AreEqual Choose(i + 1, 72, 105, 32, 77, 111, 109, 32, 45, 43, 74, 106, 111, 45, 45, 43, 65, 67, 69, 45), b(i)
    Next i
End Sub

Public Sub TestJapaneseWord_nihongo()
    ch = NewIntegers(&H65E5, &H672C, &H8A9E)
    b = e.GetBytes(ch)
    Assert.AreEqual 10, CorArray.GetLength(b)
    Dim i As Long
    For i = 0 To 9
        Assert.AreEqual Choose(i + 1, 43, 90, 101, 86, 110, 76, 73, 113, 101, 45), b(i)
    Next i
End Sub

Public Sub TestBase64AndDirect()
    ch = NewIntegers(&H65E5, 65)
    b = e.GetBytes(ch)
    Assert.AreEqual 6, CorArray.GetLength(b)
    Dim i As Long
    For i = 0 To 5
        Assert.AreEqual Choose(i + 1, 43, 90, 101, 85, 45, 65), b(i)
    Next i
End Sub

Public Sub TestNonDirectAndDirect()
    ch = NewIntegers(33, 65)
    b = e.GetBytes(ch)
    Assert.AreEqual 6, CorArray.GetLength(b)
    Dim i As Long
    For i = 0 To 5
        Assert.AreEqual Choose(i + 1, 43, 65, 67, 69, 45, 65), b(i)
    Next i
End Sub

Public Sub TestComplexCombo()
    b = e.GetBytes(ChrW$(&H65E5) & "hi!" & ChrW$(&H8A9E))
    Dim i As Long
    For i = 0 To 14
        Assert.AreEqual Choose(i + 1, 43, 90, 101, 85, 45, 104, 105, 43, 65, 67, 71, 75, 110, 103, 45), b(i)
    Next i
End Sub

Public Sub TestEncoderWithFlush()
    Dim c As Encoder
    Dim i As Long
    
    Set c = e.GetEncoder
    ch = NewIntegers(&H65E5)
    ReDim b(10)
    
    Assert.AreEqual 5, c.GetBytes(ch, 0, 1, b, 0, True)
    For i = 0 To 4
        Assert.AreEqual Choose(i + 1, 43, 90, 101, 85, 45), b(i)
    Next i
End Sub

Public Sub TestEncoderWithoutFlush()
    Dim c As Encoder
    Dim i As Long
    
    Set c = e.GetEncoder
    ReDim ch(0)
    ch(0) = &H65E5
    ReDim b(3)
    
    Assert.AreEqual 3, c.GetBytes(ch, 0, 1, b, 0, False)
    For i = 0 To 2
        Assert.AreEqual Choose(i + 1, 43, 90, 101), b(i)
    Next i
    
    ch(0) = 97
    Assert.AreEqual 3, c.GetBytes(ch, 0, 1, b, 0, True)
    For i = 0 To 2
        Assert.AreEqual Choose(i + 1, 85, 45, 97), b(i)
    Next i
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
    
    Assert.AreEqual 2565, CorArray.GetLength(b)

    FF = FreeFile
    Open App.Path & "\UTF7_Chars_0to1000.txt" For Binary As #FF
    ReDim fileBytes(0 To LOF(FF) - 1)
    Get #FF, , fileBytes
    Close #FF

    For i = 0 To CorArray.GetLength(b) - 1
        Assert.AreEqual fileBytes(i), b(i)
    Next i
End Sub

Public Sub TestEncodeEmptyString()
    Dim b() As Byte
    
    b = e.GetBytes("")
    Assert.IsFalse CorArray.IsNull(b), "Array should not be null."
    Assert.AreEqual 0, CorArray.GetLength(b)
    
    ReDim b(10)
    Assert.AreEqual 0, e.GetBytesEx("", 0, 0, b, 0)
End Sub

Public Sub TestEncodeDecode0to255()
    Dim i As Long
    Dim b() As Byte

    ReDim b(255)
    For i = 0 To 255
        b(i) = i
    Next i

    Dim s As String
    s = Encoding.UTF7.GetString(b)
    
    Erase b
    b = Encoding.UTF7.GetBytes(s)
    Assert.AreEqual 575, CorArray.GetLength(b)
    
End Sub
