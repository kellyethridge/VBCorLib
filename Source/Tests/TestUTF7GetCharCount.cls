VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF7GetCharCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF7GetCharCount
'
Option Explicit
Implements ICategorizable

Private e As UTF7Encoding
Private b() As Byte

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub Setup()
    Set e = New UTF7Encoding
End Sub

Public Sub TestGetPreamble()
    Dim b() As Byte
    
    b = e.GetPreamble
    Assert.IsFalse CorArray.IsNull(b)
    Assert.AreEqual 0, CorArray.Length(b)
End Sub

Public Sub TestPlus()
    b = NewBytes(43)
    Assert.AreEqual 0, e.GetCharCount(b)
End Sub

Public Sub TestOneNonDirect()
    b = NewBytes(33)
    Assert.AreEqual 1, e.GetCharCount(b)
End Sub

Public Sub TestPlusMinus()
    b = NewBytes(43, 45)
    Assert.AreEqual 1, e.GetCharCount(b)
End Sub

Public Sub TestSimple()
    b = e.GetBytes("hello")
    Assert.AreEqual 5, e.GetCharCount(b)
End Sub

Public Sub TestJapaneseWord_nihongo()
    Dim ch() As Integer
    ch = NewIntegers(&H65E5, &H672C, &H8A9E)
    b = e.GetBytes(ch)
    Assert.AreEqual 3, e.GetCharCount(b)
End Sub

Public Sub TestCombo()
    b = e.GetBytes(ChrW$(&H65E5) & "hi!" & ChrW$(&H8A9E))
    Assert.AreEqual 5, e.GetCharCount(b)
End Sub

Public Sub TestDecoder()
    Dim ch(10) As Integer
    Dim i As Long
    Dim c As Decoder
    
    Set c = e.GetDecoder
    
    b(0) = 43
    Assert.AreEqual 0, c.GetCharCount(b, 0, 1)
    b(0) = 33
    Assert.AreEqual 1, c.GetCharCount(b, 0, 1)
    
    b = e.GetBytes(ChrW$(&H6535))
    Assert.AreEqual 0, c.GetCharCount(b, 0, 1)
    Assert.AreEqual 0, c.GetCharCount(b, 0, 2)
    Assert.AreEqual 0, c.GetCharCount(b, 0, 3)
    Assert.AreEqual 1, c.GetCharCount(b, 0, 4)
    Assert.AreEqual 1, c.GetCharCount(b, 0, 5)
End Sub

Public Sub TestEmptyArray()
    Dim b() As Byte
    
    b = CorArray.CreateInstance(vbByte)
    Assert.AreEqual 0, e.GetCharCount(b)
End Sub
