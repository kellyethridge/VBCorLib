VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUnicodeEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUnicodeEncoding
'
Option Explicit
Implements ICategorizable

Private En As UnicodeEncoding

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub


Public Sub Setup()
    Set En = NewUnicodeEncoding(False, True)
End Sub

Public Sub TestByteOrderMark()
    Dim b() As Byte
    
    b = En.GetPreamble
    Assert.AreEqual 2, CorArray.Length(b)
    Assert.AreEqual &HFF, b(0)
    Assert.AreEqual &HFE, b(1)
End Sub

Public Sub TestGetByteCount()
    Assert.AreEqual 100, En.GetByteCount(String$(50, " "))
End Sub

Public Sub TestGetCharCount()
    Dim b() As Byte
    
    b = CorArray.CreateInstance(vbByte, 100)
    Assert.AreEqual 50, En.GetCharCount(b)
End Sub

Public Sub TestGetBytes()
    Dim b() As Byte
    Dim i As Long
    
    b = En.GetBytes(String$(50, " "))
    Assert.AreEqual 100, CorArray.Length(b)
    
    For i = 0 To UBound(b)
        If i And 1 Then
            Assert.AreEqual 0, b(i)
        Else
            Assert.AreEqual 32, b(i)
        End If
    Next i
End Sub

Public Sub TestGetChars()
    Dim ch() As Integer
    Dim b() As Byte
    Dim i As Long
    
    b = En.GetBytes(String$(50, " "))
    ch = En.GetChars(b)
    Assert.AreEqual 50, CorArray.Length(ch)
    
    For i = 0 To UBound(ch)
        Assert.AreEqual 32, ch(i)
    Next i
End Sub

Public Sub TestDecoderWithEvenBytes()
    Dim c As Decoder
    Dim b(99) As Byte
    Dim ch(100) As Integer
    Dim i As Long
    
    For i = 0 To 99
        If (i And 1) = 0 Then
            b(i) = 32
        End If
    Next i
    
    Set c = En.GetDecoder
    Assert.AreEqual 50, c.GetCharCount(b, 0, 100)
    Assert.AreEqual 50, c.GetChars(b, 0, 100, ch, 0)
    
    For i = 0 To 49
        Assert.AreEqual 32, ch(i)
    Next i
End Sub

Public Sub TestDecoderWithOddBytes()
    Dim c As Decoder
    Dim b(99) As Byte
    Dim ch(100) As Integer
    Dim i As Long
    
    For i = 0 To 99
        If (i And 1) = 0 Then
            b(i) = 32
        End If
    Next i
    
    Set c = En.GetDecoder
    Assert.AreEqual 18, c.GetCharCount(b, 0, 37)
    Assert.AreEqual 18, c.GetChars(b, 0, 37, ch, 0)
    
    For i = 0 To 17
        Assert.AreEqual 32, ch(i)
    Next i
    
    ch(0) = 0
    Assert.AreEqual 1, c.GetCharCount(b, 37, 1)
    Assert.AreEqual 1, c.GetChars(b, 37, 1, ch, 0)
    Assert.AreEqual 32, ch(0)
End Sub

Public Sub TestGetString()
    Dim b() As Byte
    
    b = En.GetBytes("Hello, World!")
    Assert.AreEqual "Hello, World!", En.GetString(b)
End Sub

Public Sub TestEncodeEmptyString()
    Dim b() As Byte
    
    b = En.GetBytes("")
    Assert.IsFalse CorArray.IsNull(b), "Array should not be null."
    Assert.AreEqual 0, CorArray.Length(b)
    
    ReDim b(10)
    Assert.AreEqual 0, En.GetBytesEx("", 0, 0, b, 0)
End Sub
