VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStreamWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamWriter
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub TestWriteChars0to64000()
    Dim sw As StreamWriter
    Dim i As Long
    Dim f As String
    
    f = App.Path & "\UTF8_Chars0To64000.txt"
    If Len(Dir$(f)) > 0 Then Kill f
    
    Set sw = newstreamwriter(f, NewUTF8Encoding(True))
    For i = 0 To 64000
        sw.WriteValue ChrW$(i)
    Next i
    sw.CloseWriter
    
    Dim En As Encoding
    Dim b() As Byte
    Dim FF As Integer
    Dim ch() As Integer
    
    Set En = sw.Encoding
    
    FF = FreeFile
    Open f For Binary As #FF
    ReDim b(0 To LOF(FF) - 1)
    Get #FF, , b
    Close #FF
    Kill f
    
    Assert.AreEqual 189828, cArray.GetLength(b)
    ' test for UTF8 Byte Order Mark
    Assert.AreEqual &HEF, b(0)
    Assert.AreEqual &HBB, b(1)
    Assert.AreEqual &HBF, b(2)
    
    ch = En.GetChars(b, 3, cArray.GetLength(b) - 3)
    Assert.AreEqual 64001, cArray.GetLength(ch)
    
    Dim j As Long
    For i = 0 To 64000
        AsWord(j) = ch(i)   ' we need to convert negative values to position without extending the sign
        Assert.AreEqual i, j, "Wrong character as index " & i
    Next i
    
End Sub

Public Sub TestWriteHugeStringAtOnce()
    Dim sw As StreamWriter
    Dim s As String
    Dim f As String
    Dim En As Encoding
    Dim FF As Integer
    Dim b() As Byte
    Dim i As Long
    
    f = App.Path & "\100003bytes.txt"
    If Len(Dir$(f)) > 0 Then Kill f
    
    Set sw = newstreamwriter(f, NewUTF8Encoding(True))
    
    s = String$(100000, "1")
    sw.WriteValue s
    sw.CloseWriter
    Set En = sw.Encoding
    
    FF = FreeFile
    Open f For Binary As #FF
    ReDim b(0 To LOF(FF) - 1)
    Get #FF, , b
    Close #FF
    Kill f
    Assert.AreEqual 100003, cArray.GetLength(b), "File length should be 100003 (includes Byte Order Mark)."
    
    For i = 3 To UBound(b)
        Assert.AreEqual 49, b(i)
    Next i
    
End Sub

