VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDESPaddingModes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDESPaddingModes
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub


Public Sub EncryptPartialBlockUsingPKCS7()
    Dim b() As Byte
    b = GetPartialBlock(PKCS7)
    Assert.AreEqual 8, CorArray.GetLength(b)
    
    Dim i As Long
    For i = 0 To 7
        Assert.AreEqual Choose(i + 1, 143, 101, 40, 181, 98, 60, 72, 212), b(i)
    Next i
End Sub

Public Sub EncryptPartialBlockUsingANSIX923()
    Dim b() As Byte
    b = GetPartialBlock(ANSIX923)
    Assert.AreEqual 8, CorArray.GetLength(b)
    
    Dim i As Long
    For i = 0 To 7
        Assert.AreEqual Choose(i + 1, 106, 208, 39, 52, 219, 197, 57, 88), b(i)
    Next i
End Sub

Public Sub EncryptPartialBlockUsingZeros()
    Dim b() As Byte
    b = GetPartialBlock(Zeros)
    Assert.AreEqual 8, CorArray.GetLength(b)
    
    Dim i As Long
    For i = 0 To 7
        Assert.AreEqual Choose(i + 1, 132, 77, 170, 216, 64, 36, 155, 48), b(i)
    Next i
End Sub

Public Sub EncryptPartialBlockUsingISO10126()
    Dim b1() As Byte
    Dim b2() As Byte
    
    b1 = GetPartialBlock(ISO10126)
    b2 = GetPartialBlock(ISO10126)
    
    Assert.AreEqual 8, CorArray.GetLength(b1)
    Assert.That b1, Iz.Not.EqualTo(b2)
End Sub

Public Sub EncryptPartialBlockUsingNone()
    Dim b() As Byte
    
    On Error GoTo errTrap
    b = GetPartialBlock(PaddingMode.None)
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is CryptographicException, "Wrong exception type caught."
End Sub

Public Sub EncryptFullBlockUsingPKCS7()
    Dim b() As Byte
    b = GetFullBlock(PKCS7)
    Assert.AreEqual 16, CorArray.GetLength(b)
    
    Dim i As Long
    For i = 0 To 15
        Assert.AreEqual Choose(i + 1, 161, 232, 185, 176, 207, 160, 131, 128, 207, 253, 119, 206, 250, 61, 24, 52), b(i)
    Next i
End Sub

Public Sub EncryptFullBlockUsingANSIX923()
    Dim b() As Byte
    b = GetFullBlock(ANSIX923)
    Assert.AreEqual 16, CorArray.GetLength(b)
    
    Dim i As Long
    For i = 0 To 15
        Assert.AreEqual Choose(i + 1, 161, 232, 185, 176, 207, 160, 131, 128, 147, 176, 0, 244, 240, 27, 104, 59), b(i)
    Next i
End Sub

Public Sub EncryptFullBlockUsingZeros()
    Dim b() As Byte
    b = GetFullBlock(Zeros)
    Assert.AreEqual 8, CorArray.GetLength(b)
    
    Dim i As Long
    For i = 0 To 7
        Assert.AreEqual Choose(i + 1, 161, 232, 185, 176, 207, 160, 131, 128), b(i)
    Next i
End Sub

Public Sub EncryptFullBlockUsingNone()
    Dim b() As Byte
    b = GetFullBlock(PaddingMode.None)
    Assert.AreEqual 8, CorArray.GetLength(b)
    
    Dim i As Long
    For i = 0 To 7
        Assert.AreEqual Choose(i + 1, 161, 232, 185, 176, 207, 160, 131, 128), b(i)
    Next i
End Sub

Public Sub EncryptFullBlockUsingISO10126()
    Dim b1() As Byte
    Dim b2() As Byte
    
    b1 = GetFullBlock(ISO10126)
    b2 = GetFullBlock(ISO10126)
    
    Assert.AreEqual 16, CorArray.GetLength(b1)
    
    Dim i As Long
    For i = 0 To 7
        Assert.AreEqual b1(i), b2(i)
    Next i
    
    Assert.That b1, Iz.Not.EqualTo(b2)
End Sub

Public Sub DecryptPartialBlockUsingPKCS7()
    
End Sub


Private Function GetPartialDecryptedBlock(ByRef Data() As Byte, ByVal P As PaddingMode) As Byte()
    Dim csp As New DESCryptoServiceProvider
    csp.Key = NewArray(vbByte, 1, 2, 3, 4, 5, 6, 7, 8)
    csp.IV = csp.Key
    csp.Mode = CBC
    csp.Padding = P
    
    Dim t As CryptoAPITransform
    Set t = csp.CreateDecryptor
    
    GetPartialDecryptedBlock = t.TransformFinalBlock(Data, 0, 2)
End Function

Private Function GetFullBlock(ByVal P As PaddingMode) As Byte()
    Dim csp As New DESCryptoServiceProvider
    csp.Key = NewArray(vbByte, 1, 2, 3, 4, 5, 6, 7, 8)
    csp.IV = csp.Key
    csp.Mode = CBC
    csp.Padding = P
    
    Dim Data() As Byte
    Data = NewArray(vbByte, 65, 66, 67, 68, 69, 70, 71, 72)
    
    Dim t As CryptoAPITransform
    Set t = csp.CreateEncryptor
    
    GetFullBlock = t.TransformFinalBlock(Data, 0, 8)
End Function

Private Function GetPartialBlock(ByVal P As PaddingMode) As Byte()
    Dim csp As New DESCryptoServiceProvider
    csp.Key = NewArray(vbByte, 1, 2, 3, 4, 5, 6, 7, 8)
    csp.IV = csp.Key
    csp.Mode = CBC
    csp.Padding = P
    
    Dim Data() As Byte
    Data = NewArray(vbByte, 65, 66)
    
    Dim t As CryptoAPITransform
    Set t = csp.CreateEncryptor
    
    GetPartialBlock = t.TransformFinalBlock(Data, 0, 2)
End Function
