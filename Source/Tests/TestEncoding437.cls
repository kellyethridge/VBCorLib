VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestEncoding437"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestEncoding437
'
Option Explicit

Private en As Encoding

Public Sub Setup()
    Set en = Encoding.GetEncoding(437)
End Sub

Public Sub TestGetByteCount()
    Assert.AreEqual 65536, en.GetByteCount(GetChars), "ByteCount should return 65535."
End Sub

Public Sub TestGetBytes()
    Dim Expected() As Byte
    Dim Actual() As Byte
    Dim i As Long
    
    Expected = GetEncodedBytes
    
    Actual = en.GetBytes(GetChars)
    
    For i = 0 To UBound(Actual)
        Assert.AreEqual Expected(i), Actual(i)
    Next i
End Sub

Public Sub TestGetCharCount()
    Dim Expected() As Byte
    
    Expected = GetEncodedBytes
    Assert.AreEqual 65536, en.GetCharCount(Expected)
End Sub

Public Sub TestGetBytesFromString()
    Dim b() As Byte
    
    b = en.GetBytes("Hello")
    Assert.AreEqual 4, UBound(b), "Upper bound should be 4."
End Sub



Private Function GetChars() As String
    Dim sb As StringBuilder
    Dim i As Long
    
    Set sb = New StringBuilder
    
    For i = 0 To 65535
        sb.AppendChar i
    Next i
    
    GetChars = sb.ToString
End Function

Private Function GetEncodedBytes() As Byte()
    Dim f As Long
    Dim ret() As Byte
    
    f = FreeFile
    Open App.Path & "\Dummy Test Files\Chars0To65535InCodePage437.txt" For Binary As #f
    ReDim ret(0 To LOF(f) - 1)
    Get #f, , ret
    Close #f
    
    GetEncodedBytes = ret
End Function

Public Sub TestEncodeEmptyString()
    Dim b() As Byte
    
    b = en.GetBytes("")
    Assert.IsFalse cArray.IsNull(b)
    Assert.AreEqual 0, cArray.GetLength(b)
    
    ReDim b(10)
    Assert.AreEqual 0, en.GetBytesEx("", 0, 0, b, 0)
End Sub

Public Sub TestDecodeEmptyArray()
    Dim b() As Byte
    Dim ch() As Integer
    
    b = cArray.CreateInstance(vbByte)
    ch = en.GetChars(b)
    Assert.IsFalse cArray.IsNull(ch), "Array should not be null."
    Assert.AreEqual 0, cArray.GetLength(ch)
    
    ReDim ch(10)
    Assert.AreEqual 0, en.GetCharsEx(b, 0, 0, ch, 0)

End Sub
