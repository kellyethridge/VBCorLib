VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamWriterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamWriter
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "WriteValue_WithValues_WritesExpected"
            Test.Use CByte(1), NewBytes(49)
            Test.Use CInt(1), NewBytes(49)
            Test.Use CLng(1), NewBytes(49)
            Test.Use CLng(12), NewBytes(49, 50)
            Test.Use CLng(12345678), NewBytes(49, 50, 51, 52, 53, 54, 55, 56)
            Test.Use 1#, NewBytes(49)
            Test.Use 1.1, NewBytes(49, 46, 49)
            Test.Use CDec(1.1), NewBytes(49, 46, 49)
            Test.Use "Hello", NewBytes(72, 101, 108, 108, 111)
            
            Dim Obj As New Int32
            Obj.Value = 10
            Test.Use Obj, NewBytes(49, 48)
            
    End Select
End Sub

Public Sub NewStreamWriter_WithNothingSource_ThrowsArgumentNullException()
    On Error Resume Next
    
    Cor.NewStreamWriter Nothing
    
    AssertArgumentNullException Err, "Source"
End Sub

Public Sub NewStreamWriter_WithNonStreamObject_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewStreamWriter Me
     
    AssertArgumentException Err, "Source"
End Sub

Public Sub NewStreamWriter_WithNegativeBufferSize_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Cor.NewStreamWriter New MemoryStream, BufferSize:=-1
    
    AssertArgumentOutOfRangeException Err, "BufferSize"
End Sub

Public Sub NewStreamWriter_WithNoEncoding_DefaultsToUTF8Encoding()
    Dim Actual As Encoding
    
    With Cor.NewStreamWriter(New MemoryStream)
        Set Actual = .Encoding
    End With
    
    Assert.That TypeOf Actual Is UTF8Encoding, Iz.True
End Sub

Public Sub WriteValue_WithValues_WritesExpected(ByRef Value As Variant, ByRef Expected() As Byte)
    Dim Store As New MemoryStream
        
    With NewStreamWriter(Store)
        .WriteValue Value
    End With
    
    Dim Actual() As Byte
    Actual = Store.ToArray
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteChars_WithNullChars_ThrowsArgumentNullException()
    Dim NullChars() As Integer
    On Error Resume Next
    
    With Cor.NewStreamWriter(New MemoryStream)
        .WriteChars NullChars
    End With
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub WriteChars_WithChars_WritesToStream()
    Dim Store   As New MemoryStream
    Dim Chars() As Integer
    Chars = NewIntegers(48, 49, 50)
    
    With Cor.NewStreamWriter(Store)
        .WriteChars Chars
    End With
    
    Assert.That Store.ToArray, Iz.EqualTo(NewBytes(48, 49, 50))
End Sub

'Public Sub TestWriteChars0to64000()
'    Dim sw As StreamWriter
'    Dim i As Long
'    Dim f As String
'
'    f = App.Path & "\UTF8_Chars0To64000.txt"
'    If Len(Dir$(f)) > 0 Then Kill f
'
'    Set sw = newstreamwriter(f, NewUTF8Encoding(True))
'    For i = 0 To 64000
'        sw.WriteValue ChrW$(i)
'    Next i
'    sw.CloseWriter
'
'    Dim En As Encoding
'    Dim b() As Byte
'    Dim FF As Integer
'    Dim ch() As Integer
'
'    Set En = sw.Encoding
'
'    FF = FreeFile
'    Open f For Binary As #FF
'    ReDim b(0 To LOF(FF) - 1)
'    Get #FF, , b
'    Close #FF
'    Kill f
'
'    Assert.AreEqual 189828, cArray.GetLength(b)
'    ' test for UTF8 Byte Order Mark
'    Assert.AreEqual &HEF, b(0)
'    Assert.AreEqual &HBB, b(1)
'    Assert.AreEqual &HBF, b(2)
'
'    ch = En.GetChars(b, 3, cArray.GetLength(b) - 3)
'    Assert.AreEqual 64001, cArray.GetLength(ch)
'
'    Dim j As Long
'    For i = 0 To 64000
'        AsWord(j) = ch(i)   ' we need to convert negative values to position without extending the sign
'        Assert.AreEqual i, j, "Wrong character as index " & i
'    Next i
'
'End Sub
'
'Public Sub TestWriteHugeStringAtOnce()
'    Dim sw As StreamWriter
'    Dim s As String
'    Dim f As String
'    Dim En As Encoding
'    Dim FF As Integer
'    Dim b() As Byte
'    Dim i As Long
'
'    f = App.Path & "\100003bytes.txt"
'    If Len(Dir$(f)) > 0 Then Kill f
'
'    Set sw = newstreamwriter(f, NewUTF8Encoding(True))
'
'    s = String$(100000, "1")
'    sw.WriteValue s
'    sw.CloseWriter
'    Set En = sw.Encoding
'
'    FF = FreeFile
'    Open f For Binary As #FF
'    ReDim b(0 To LOF(FF) - 1)
'    Get #FF, , b
'    Close #FF
'    Kill f
'    Assert.AreEqual 100003, cArray.GetLength(b), "File length should be 100003 (includes Byte Order Mark)."
'
'    For i = 3 To UBound(b)
'        Assert.AreEqual 49, b(i)
'    Next i
'
'End Sub
'

