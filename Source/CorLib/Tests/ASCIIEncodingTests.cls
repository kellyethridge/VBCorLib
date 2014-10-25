VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASCIIEncodingTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestASCIIEncoding
'
Option Explicit
Implements ITestCaseSource

Private en As Encoding

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetBytes_WithCharacters_ReturnsExpected"
            Test.Use Chr$(0), NewBytes(0)
            
    
    End Select
End Sub

Public Sub Setup()
    Set en = Encoding.ASCII
End Sub

Public Sub GetBytes_WithCharacters_ReturnsExpected(ByVal StringToEncode As String, ByRef Expected() As Byte)
    Dim Encoding As New ASCIIEncoding
    
    Dim Actual() As Byte
    Actual = Encoding.GetBytes(StringToEncode)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TestGetBytesFromInvalidChars()
    Dim b() As Byte
    Dim s As String
    Dim i As Long
    
    For i = 0 To 1000
        s = s & ChrW$(i)
    Next i
    
    b = en.GetBytes(s)
    
    Assert.AreEqual 1000, UBound(b)
    
    For i = 0 To 1000
        If i < 128 Then
            Assert.AreEqual i, b(i)
        Else
            Assert.AreEqual 63, b(i)
        End If
    Next i
    
End Sub

Public Sub TestEmptyString()
    Dim b() As Byte
    
    b = en.GetBytes("")
    Assert.IsFalse cArray.IsNull(b), "Array should not be null."
    Assert.AreEqual 0, cArray.GetLength(b)
    
    ReDim b(10)
    Assert.AreEqual 0, en.GetBytesEx("", 0, 0, b, 0)
End Sub

Public Sub TestEmptyArray()
    Dim b() As Byte
    Dim ch() As Integer
    
    b = cArray.CreateInstance(vbByte)
    ch = en.GetChars(b)
    Assert.IsFalse cArray.IsNull(ch), "Array should not be null."
    Assert.AreEqual 0, cArray.GetLength(ch)
    
    ReDim ch(10)
    Assert.AreEqual 0, en.GetCharsEx(b, 0, 0, ch, 0)
    
End Sub

Public Sub TestGetString()
    Dim b(10) As Byte
    Dim i As Long
    
    For i = 0 To 9
        b(i) = 48 + i
    Next i
    b(10) = 237
    
    Assert.AreEqual "0123456789?", en.GetString(b)
End Sub

Public Sub TestGetStringWords()
    Dim b() As Byte
    b = Cor.NewBytes(&H48, &H65, &H6C, &H6C, &H6F, &H20, &H74, &H68, &H65, &H72, &H65)
    
    Dim s As String
    s = en.GetString(b)
    Assert.AreEqual "Hello there", s
End Sub

Public Sub TestGetStringFromPartOfArray()
    Dim b() As Byte
    b = StrConv("Now is the time for", vbFromUnicode)
    
    Assert.AreEqual "Now is t", en.GetString(b, 0, 8)
End Sub


