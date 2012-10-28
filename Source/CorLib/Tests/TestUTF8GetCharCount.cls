VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF8GetCharCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   TestUTF8GetCharCount
'
Option Explicit

Private e As UTF8Encoding
Private b() As Byte

Public Sub Setup()
    Set e = NewUTF8Encoding(False, True)
End Sub

Public Sub TestAllAscii()
    b = cArray.NewArray(ciByte, 85, 84, 70, 56, 32, 69, 110, 99, 111, 100, 105, 110, 103, 32, 69, 120, 97, 109, 112, 108, 101)
    
    Assert.AreEqual 21, e.GetCharCount(b)
    Assert.AreEqual 8, e.GetCharCount(b, 2, 8)
End Sub

Public Sub Test2ByteChars()
    b = e.GetBytes(cArray.NewArray(ciInteger, &HD801, &HDC01))
    Assert.AreEqual 2, e.GetCharCount(b)
End Sub

Public Sub TestNonAscii()
    b = e.GetBytes(String$(10, ChrW$(150)))
    Assert.AreEqual 10, e.GetCharCount(b)
End Sub

Public Sub TestHighChars()
    b = e.GetBytes(String$(10, ChrW$(&H2030)))
    Assert.AreEqual 10, e.GetCharCount(b)
End Sub

Public Sub TestComboChars()
    Dim asciiChars As String
    Dim nonAsciiChars As String
    Dim surrogateChars As String
    Dim i As Long
    
    asciiChars = String$(10, "A")
    nonAsciiChars = String$(10, ChrW$(167))
    For i = 1 To 10
        surrogateChars = surrogateChars & ChrW$(&HD801) & ChrW$(&HDC01)
    Next i
    Dim s As String
    s = nonAsciiChars & surrogateChars & asciiChars
    
    b = e.GetBytes(s)
    Assert.AreEqual 40, e.GetCharCount(b)
End Sub


Public Sub TestOverLong1()
    b = cArray.NewArray(ciByte, &HC0, &H8A)
    On Error GoTo errTrap
    e.GetCharCount b
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex), "An exception should be caught."
End Sub

Public Sub TestOverLong2()
    b = cArray.NewArray(ciByte, &HE0, &H80, &H8A)
    On Error GoTo errTrap
    e.GetCharCount b
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."

errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex), "An exception should be caught."
End Sub

Public Sub TestOverLong3()
    b = cArray.NewArray(ciByte, &HF0, &H80, &H80, &H8A)
    On Error GoTo errTrap
    e.GetCharCount b
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."

errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is ArgumentException
End Sub

Public Sub TestZeroLengthArray()
    b = cArray.CreateInstance(ciByte)
    Assert.AreEqual 0, e.GetCharCount(b)
End Sub
