VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTripleDESEncryption128bit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTripleDESEncryption
'
Option Explicit

Private csp As TripleDESCryptoServiceProvider
Private trans As CryptoAPITransform

Public Sub Setup()
    Set csp = New TripleDESCryptoServiceProvider
    csp.Key = NewArray(ciByte, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 1, 2, 3, 4, 5, 6)
    csp.IV = NewArray(ciByte, 1, 2, 3, 4, 5, 6, 7, 8)
    csp.Padding = PKCS7
    csp.Mode = CBC
    Set trans = csp.CreateEncryptor
End Sub

Public Sub EncryptOneByte128BitKey()
    Dim b() As Byte
    b = NewArray(ciByte, 1)
    
    Dim t As ICryptoTransform
    Set t = csp.CreateEncryptor
    
    Dim b2() As Byte
    b2 = t.TransformFinalBlock(b, 0, 1)
    Assert.AreEqual 8, cArray.GetLength(b2)
    
    Dim i As Long
    For i = 1 To 8
        Assert.AreEqual CLng(Choose(i, 78, 228, 36, 62, 169, 116, 61, 240)), b2(i - 1)
    Next i
    
    ' repeat to test resetting is correct
    Erase b2
    b2 = t.TransformFinalBlock(b, 0, cArray.GetLength(b))
    Assert.AreEqual 8, cArray.GetLength(b2)
    For i = 1 To 8
        Assert.AreEqual CLng(Choose(i, 78, 228, 36, 62, 169, 116, 61, 240)), b2(i - 1)
    Next i
End Sub

Public Sub EncryptFirst8Bytes()
    Dim src() As Byte
    src = Encoding.ASCII.GetBytes("Now is the time for")

    Dim dst(100) As Byte

    Dim t As ICryptoTransform
    Set t = csp.CreateEncryptor

    Assert.AreEqual 8, t.TransformBlock(src, 0, 8, dst, 0)

    Dim i As Long
    For i = 1 To 8
        Assert.AreEqual CLng(Choose(i, 36, 7, 253, 191, 96, 146, 192, 154)), dst(i - 1)
    Next i
End Sub

