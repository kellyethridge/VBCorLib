VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF7GetByteCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF7GetByteCount
'
Option Explicit
Implements ICategorizable

Private e As UTF7Encoding

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Public Sub Setup()
    Set e = New UTF7Encoding
End Sub

Public Sub TestDirectlyEncodedCharacters()
    Assert.AreEqual 1, e.GetByteCount("a")
    Assert.AreEqual 26, e.GetByteCount(String$(26, "1"))
End Sub

Public Sub TestNonDirectUnicode()
    Assert.AreEqual 5, e.GetByteCount("!")
    Assert.AreEqual 16, e.GetByteCount(String$(5, "!"))
    Assert.AreEqual 6, e.GetByteCount("!" & vbCr)
    Assert.AreEqual 10, e.GetByteCount("+!+")
    Assert.AreEqual 4, e.GetByteCount("++")
End Sub

Public Sub TestMixture()
    Assert.AreEqual 13, e.GetByteCount("a!'=2")
End Sub

Public Sub TestPlusMinus()
    Assert.AreEqual 3, e.GetByteCount("+-")
End Sub

Public Sub TestSimpleCombo2()
    Dim ch() As Integer
    ch = NewIntegers(&H48, &H69, &H20, &H4D, &H6F, &H6D, &H20, &H2D, &H263A, &H2D, &H21)
    Assert.AreEqual 19, e.GetByteCount(ch)
End Sub

Public Sub TestJapaneseWord_nihongo()
    Dim ch() As Integer
    ch = NewIntegers(&H65E5, &H672C, &H8A9E)
    Assert.AreEqual 10, e.GetByteCount(ch)
End Sub

Public Sub TestComplexCombo()
    Assert.AreEqual 15, e.GetByteCount(ChrW$(&H65E5) & "hi!" & ChrW$(&H8A9E))
End Sub

Public Sub TestEncoder()
    Dim c As Encoder
    Dim ch(100) As Integer
    Dim i As Long
    
    For i = 0 To 100
        ch(i) = &H65E5
    Next i
    Set c = e.GetEncoder
    
    Assert.AreEqual 270, c.GetByteCount(ch, 0, 101, False)
    Assert.AreEqual 272, c.GetByteCount(ch, 0, 101, True)
    
    For i = 0 To 100
        ch(i) = IIf(i < 50, &H65E5, 97)
    Next i
    
    Assert.AreEqual 187, c.GetByteCount(ch, 0, 101, False), "Should return 187 when Flush = False."
    Assert.AreEqual 187, c.GetByteCount(ch, 0, 101, True), "Should return 187 when Flush = True."
End Sub

Public Sub TestEmptyString()
    Assert.AreEqual 0, e.GetByteCount("")
End Sub
