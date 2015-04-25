VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStringReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   TestStringReader
'
Option Explicit
Implements ICategorizable

Private sr As StringReader

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub TestRead()
    Dim ch As Long
    Dim i As Long
    
    Set sr = NewStringReader("abcdefghijk")
    i = Asc("a")
    ch = sr.Read
    Do While ch > 0
        Assert.AreEqual i, ch
        i = i + 1
        ch = sr.Read
    Loop
    Assert.AreEqual Asc("l"), i
    Assert.AreEqual -1, sr.Read
End Sub

Public Sub TestPeek()
    Set sr = NewStringReader("a")
    Assert.AreEqual Asc("a"), sr.Peek
    Assert.AreEqual Asc("a"), sr.Read
    Assert.AreEqual -1, sr.Read
End Sub

Public Sub TestReadBlock()
    Dim ch(100) As Integer
    Dim i As Long
    
    Set sr = NewStringReader("abcdefghijk")
    
    Assert.AreEqual 11, sr.ReadBlock(ch, 0, 100)
    For i = 0 To 10
        Assert.AreEqual Asc("a") + i, ch(i)
    Next i
End Sub

Public Sub TestReadBlockPastEndOfStream()
    Dim ch(10) As Integer
    Dim i As Long
    
    Set sr = Cor.NewStringReader("12345678")
    Assert.AreEqual 8, sr.ReadBlock(ch, 0, 11)
End Sub

Public Sub TestReadLine()
    Set sr = NewStringReader("abc" & vbCr & vbLf & "def" & vbLf & vbCr)
    Assert.AreEqual "abc", sr.ReadLine
    Assert.AreEqual "def", sr.ReadLine
    Assert.IsFalse cString.IsNull(sr.ReadLine)
    Assert.IsTrue cString.IsNull(sr.ReadLine)
End Sub

Public Sub TestReadToEnd()
    Set sr = NewStringReader("abcdefghi")
    sr.Read
    Assert.AreEqual "bcdefghi", sr.ReadToEnd
End Sub
