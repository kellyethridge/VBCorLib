VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestConvert
'
Option Explicit

Public Sub TestToBase64StringOneChar()
    Dim b(0) As Byte
    Dim s As String
    
    b(0) = 22
    s = Convert.ToBase64String(b)
    Assert.AreEqual "Fg==", s
End Sub

Public Sub TestToBase64String3Bytes()
    Dim b() As Byte
    Dim s As String
    
    b = NewBytes(22, 250, 42)
    s = Convert.ToBase64String(b)
    Assert.AreEqual "Fvoq", s
End Sub

Public Sub TestToBase64String3BytesOffset()
    Dim b() As Byte
    Dim s As String
    
    b = NewBytes(22, 250, 42)
    s = Convert.ToBase64String(b, 1, 2)
    Assert.AreEqual "+io=", s
End Sub

Public Sub TestToBase64String255Bytes()
    Dim b(255) As Byte
    Dim s As String
    Dim i As Long
    
    For i = 0 To 255
        b(i) = i
    Next i
    s = Convert.ToBase64String(b)
    Assert.AreEqual "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8gISIjJCUmJygpKissLS4vMDEyMzQ1Njc4OTo7PD0+P0BBQkNERUZHSElKS0xNTk9QUVJTVFVWV1hZWltcXV5fYGFiY2RlZmdoaWprbG1ub3BxcnN0dXZ3eHl6e3x9fn+AgYKDhIWGh4iJiouMjY6PkJGSk5SVlpeYmZqbnJ2en6ChoqOkpaanqKmqq6ytrq+wsbKztLW2t7i5uru8vb6/wMHCw8TFxsfIycrLzM3Oz9DR0tPU1dbX2Nna29zd3t/g4eLj5OXm5+jp6uvs7e7v8PHy8/T19vf4+fr7/P3+/w==", s
End Sub

Public Sub TestToBase64CharArraySmallBuffer()
    Dim b(0) As Byte
    Dim s As String
    Dim ch(2) As Integer
    
    b(0) = 22
    On Error GoTo errTrap
    s = Convert.ToBase64CharArray(b, 0, 1, ch, 0)
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestFromBase64CharArray()
    Dim ch() As Integer
    Dim b() As Byte
    
    ch = cString.ToCharArray("Fg==")
    b = Convert.FromBase64CharArray(ch, 0, 4)
    Assert.AreEqual 1, cArray.GetLength(b)
    Assert.AreEqual 22, b(0)
End Sub

Public Sub TestFromBase64CharArray4Chars()
    Dim ch() As Integer
    Dim b() As Byte
    
    ch = cString.ToCharArray("Fvoq")
    b = Convert.FromBase64CharArray(ch, 0, 4)
    Assert.AreEqual 3, cArray.GetLength(b)
    Assert.AreEqual 22, b(0)
    Assert.AreEqual 250, b(1)
    Assert.AreEqual 42, b(2)
End Sub

Public Sub TestFromBase64CharArrayLots()
    Dim ch() As Integer
    Dim b() As Byte
    Dim i As Long
    
    ch = cString.ToCharArray("AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8gISIjJCUmJygpKissLS4vMDEyMzQ1Njc4OTo7PD0+P0BBQkNERUZHSElKS0xNTk9QUVJTVFVWV1hZWltcXV5fYGFiY2RlZmdoaWprbG1ub3BxcnN0dXZ3eHl6e3x9fn+AgYKDhIWGh4iJiouMjY6PkJGSk5SVlpeYmZqbnJ2en6ChoqOkpaanqKmqq6ytrq+wsbKztLW2t7i5uru8vb6/wMHCw8TFxsfIycrLzM3Oz9DR0tPU1dbX2Nna29zd3t/g4eLj5OXm5+jp6uvs7e7v8PHy8/T19vf4+fr7/P3+/w==")
    b = Convert.FromBase64CharArray(ch, 0, cArray.GetLength(ch))
    
    Assert.AreEqual 256, cArray.GetLength(b)
    
    For i = 0 To 255
        Assert.AreEqual i, b(i)
    Next i
End Sub

Public Sub TestFromBase64String()
    Dim b() As Byte
    Dim i As Long
    
    b = Convert.FromBase64String("AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8gISIjJCUmJygpKissLS4vMDEyMzQ1Njc4OTo7PD0+P0BBQkNERUZHSElKS0xNTk9QUVJTVFVWV1hZWltcXV5fYGFiY2RlZmdoaWprbG1ub3BxcnN0dXZ3eHl6e3x9fn+AgYKDhIWGh4iJiouMjY6PkJGSk5SVlpeYmZqbnJ2en6ChoqOkpaanqKmqq6ytrq+wsbKztLW2t7i5uru8vb6/wMHCw8TFxsfIycrLzM3Oz9DR0tPU1dbX2Nna29zd3t/g4eLj5OXm5+jp6uvs7e7v8PHy8/T19vf4+fr7/P3+/w==")
    
    Assert.AreEqual 256, cArray.GetLength(b)
    
    For i = 0 To 255
        Assert.AreEqual i, b(i)
    Next i
End Sub

Public Sub TestToFromBase64StringRange0to300()
    Dim b() As Byte
    Dim b2() As Byte
    Dim i As Long
    Dim j As Long
    Dim s As String
    
    For i = 0 To 300
        ReDim b(i)
        For j = UBound(b) To LBound(b) Step -1
            b(j) = j And &HFF
        Next j
        s = Convert.ToBase64String(b)
        b2 = Convert.FromBase64String(s)
        Assert.AreEqual cArray.GetLength(b), cArray.GetLength(b2)
        For j = UBound(b) To LBound(b) Step -1
            Assert.AreEqual b(j), b2(j)
        Next j
    Next i
End Sub

Public Sub TestToBase64StringLineBreaks()
    Dim i As Long
    Dim s As String
    Dim b() As Byte
    
    ReDim b(255)
    For i = 0 To 255
        b(i) = i
    Next i
    
    s = Convert.ToBase64String(b, , , True)
    Assert.AreEqual 352, Len(s)
    
    i = InStr(s, vbCrLf)
    Assert.AreEqual 77, i
    i = InStr(i + 2, s, vbCrLf)
    Assert.AreEqual 78 + 77, i
    i = InStr(i + 2, s, vbCrLf)
    Assert.AreEqual 78 * 2 + 77, i
    i = InStr(i + 2, s, vbCrLf)
    Assert.AreEqual 78 * 3 + 77, i
    Assert.AreEqual 351, InStr(s, "==")
End Sub

Public Sub TestFromBase64StringLineBreaks()
    Dim i As Long
    Dim s As String
    Dim b() As Byte
    
    ReDim b(255)
    For i = 0 To 255
        b(i) = i
    Next i
    
    s = Convert.ToBase64String(b, , , True)
        
    b = Convert.FromBase64String(s)
    
    Assert.AreEqual 256, cArray.GetLength(b)
    For i = 0 To 255
        Assert.AreEqual i, b(i)
    Next i
End Sub
