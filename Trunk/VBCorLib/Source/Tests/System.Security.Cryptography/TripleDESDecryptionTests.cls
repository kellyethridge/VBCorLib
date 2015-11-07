VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TripleDESDecryptionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TripleDESDecryptionTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "TransformFinalBlock_WithCipherMode_ReturnsExpected"
            AddTests Test
            
        Case "TransformBlock_WithCipherMode_ReturnsExpected"
            AddTests Test

        Case "TransformFinalBlock_WithPaddingMode_ReturnsExpected"
'            Test.Use PaddingMode.None, "05c9c4cafb9937d95bbfbedfc5d77f19a6cd5a5dab188a33dfd8979fc4b7b2be"
            Test.Use PaddingMode.ANSIX923, "8c5c1e4affcaf69cc8c274a95f57c4a41b63c78e2ae3c67e6192c492157fac73623f2d96ae4ecfa6bc7e61ee2557ed205bc894a5c4938001"
            Test.Use PaddingMode.PKCS7, "8c5c1e4affcaf69cc8c274a95f57c4a41b63c78e2ae3c67e6192c492157fac73623f2d96ae4ecfa6bc7e61ee2557ed202193aad21d63874c"
            
    End Select
End Sub

Private Sub AddTests(ByVal Test As TestCaseCollector)
    Test.Use CipherMode.ECB, "12345678abcdefghstuvwxyz", "00000000", "8c5c1e4affcaf69cc8c274a95f57c4a41b63c78e2ae3c67e6192c492157fac73623f2d96ae4ecfa6bc7e61ee2557ed203f144d2cc8e8e10e"
    Test.Use CipherMode.CBC, "12345678abcdefghstuvwxyz", "abcdefgh", "61b0cefb60b56d1885fcf647d7ebf44c9031b2f2c2c06018d749f99620abe1b226ce32411e55d2ba23b8704f6e19de22461d8f2bf1cc82cf"
End Sub

Public Sub TransformFinalBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByRef Crypted As String)
    Dim csp         As New TripleDESCryptoServiceProvider
    Dim Decryptor   As ICryptoTransform
    Dim MCrypt()    As Byte
    csp.Mode = CipherMode
    csp.Padding = PaddingMode.None
    Set Decryptor = csp.CreateDecryptor(Encoding.ASCII.GetBytes(Key), Encoding.ASCII.GetBytes(IV))
    MCrypt = HexToBytes(Crypted)
    
    Dim Actual() As Byte
    Actual = Decryptor.TransformFinalBlock(MCrypt, 0, CorArray.Length(MCrypt))
    
    Const Expected As String = "This is the message to encrypt, it will be padded."
    Assert.That CorString.TrimEnd(Encoding.ASCII.GetString(Actual), vbNullChar), Iz.EqualTo(Expected)
End Sub

Public Sub TransformBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByRef Crypted As String)
    Dim csp         As New TripleDESCryptoServiceProvider
    Dim Decryptor   As ICryptoTransform
    Dim MCrypt()    As Byte
    csp.Mode = CipherMode
    csp.Padding = PaddingMode.None
    Set Decryptor = csp.CreateDecryptor(Encoding.ASCII.GetBytes(Key), Encoding.ASCII.GetBytes(IV))
    MCrypt = HexToBytes(Crypted)

    Dim ms              As New MemoryStream
    Dim Actual          As Long
    Dim ActualOutput()  As Byte
    ReDim ActualOutput(UBound(MCrypt))
    Actual = Decryptor.TransformBlock(MCrypt, 0, 48, ActualOutput, 0)
    ms.WriteBlock ActualOutput, 0, 48
    ms.WriteBlock Decryptor.TransformFinalBlock(MCrypt, 48, 8), 0, 8
    ActualOutput = ms.ToArray

    Const Expected As String = "This is the message to encrypt, it will be padded."
    Assert.That Actual, Iz.EqualTo(48)
    Assert.That CorString.TrimEnd(Encoding.ASCII.GetString(ActualOutput), vbNullChar), Iz.EqualTo(Expected)
End Sub

Public Sub TransformFinalBlock_WithPaddingMode_ReturnsExpected(ByVal Padding As PaddingMode, ByRef Crypted As String)
    Dim csp         As New TripleDESCryptoServiceProvider
    Dim Decryptor   As ICryptoTransform
    Dim MCrypt()    As Byte
    csp.Mode = CipherMode.ECB
    csp.Padding = Padding
    Set Decryptor = csp.CreateDecryptor(Encoding.ASCII.GetBytes("12345678abcdefghstuvwxyz"), Encoding.ASCII.GetBytes("00000000"))
    MCrypt = HexToBytes(Crypted)

    Dim Actual() As Byte
    Actual = Decryptor.TransformFinalBlock(MCrypt, 0, CorArray.Length(MCrypt))

    Const Expected As String = "This is the message to encrypt, it will be padded."
    Assert.That Encoding.ASCII.GetString(Actual), Iz.EqualTo(Expected)
End Sub

