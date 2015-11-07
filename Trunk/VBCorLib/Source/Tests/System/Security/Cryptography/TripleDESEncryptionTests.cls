VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TripleDESEncryptionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TripleDESEncryptionTests
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
            Test.Use PaddingMode.Zeros, "8c5c1e4affcaf69cc8c274a95f57c4a41b63c78e2ae3c67e6192c492157fac73623f2d96ae4ecfa6bc7e61ee2557ed203f144d2cc8e8e10e"
            Test.Use PaddingMode.ANSIX923, "8c5c1e4affcaf69cc8c274a95f57c4a41b63c78e2ae3c67e6192c492157fac73623f2d96ae4ecfa6bc7e61ee2557ed205bc894a5c4938001"
            Test.Use PaddingMode.PKCS7, "8c5c1e4affcaf69cc8c274a95f57c4a41b63c78e2ae3c67e6192c492157fac73623f2d96ae4ecfa6bc7e61ee2557ed202193aad21d63874c"
            
    End Select
End Sub

Private Sub AddTests(ByVal Test As TestCaseCollector)
    Test.Use CipherMode.ECB, "12345678abcdefghstuvwxyz", "00000000", "8c5c1e4affcaf69cc8c274a95f57c4a41b63c78e2ae3c67e6192c492157fac73623f2d96ae4ecfa6bc7e61ee2557ed203f144d2cc8e8e10e"
    Test.Use CipherMode.CBC, "12345678abcdefghstuvwxyz", "abcdefgh", "61b0cefb60b56d1885fcf647d7ebf44c9031b2f2c2c06018d749f99620abe1b226ce32411e55d2ba23b8704f6e19de22461d8f2bf1cc82cf"
End Sub

Public Sub TransformFinalBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByRef Expected As String)
    Dim csp         As New TripleDESCryptoServiceProvider
    Dim Encryptor   As ICryptoTransform
    Dim Message()   As Byte
    csp.Mode = CipherMode
    csp.Padding = PaddingMode.Zeros
    Set Encryptor = csp.CreateEncryptor(Encoding.ASCII.GetBytes(Key), Encoding.ASCII.GetBytes(IV))
    Message = Encoding.ASCII.GetBytes("This is the message to encrypt, it will be padded.")

    Dim Actual() As Byte
    Actual = Encryptor.TransformFinalBlock(Message, 0, CorArray.Length(Message))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub TransformBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByRef Expected As String)
    Dim csp         As New TripleDESCryptoServiceProvider
    Dim Encryptor   As ICryptoTransform
    Dim Message()   As Byte
    csp.Mode = CipherMode
    csp.Padding = PaddingMode.Zeros
    Set Encryptor = csp.CreateEncryptor(Encoding.ASCII.GetBytes(Key), Encoding.ASCII.GetBytes(IV))
    Message = Encoding.ASCII.GetBytes("This is the message to encrypt, it will be padded.")

    Dim ms              As New MemoryStream
    Dim Actual          As Long
    Dim ActualOutput()  As Byte
    ReDim ActualOutput(UBound(Message))
    Actual = Encryptor.TransformBlock(Message, 0, 48, ActualOutput, 0)
    ms.WriteBlock ActualOutput, 0, 48
    ms.WriteBlock Encryptor.TransformFinalBlock(Message, 48, 2), 0, 8
    ActualOutput = ms.ToArray
    
    Assert.That Actual, Iz.EqualTo(48)
    Assert.That ActualOutput, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub TransformFinalBlock_WithPaddingMode_ReturnsExpected(ByVal Mode As PaddingMode, ByRef Expected As String)
    Dim csp As New TripleDESCryptoServiceProvider
    Dim Encryptor As ICryptoTransform
    Dim Message() As Byte
    csp.Mode = CipherMode.ECB
    csp.Padding = Mode
    Set Encryptor = csp.CreateEncryptor(Encoding.ASCII.GetBytes("12345678abcdefghstuvwxyz"), Encoding.ASCII.GetBytes("00000000"))
    Message = Encoding.ASCII.GetBytes("This is the message to encrypt, it will be padded.")

    Dim Actual() As Byte
    Actual = Encryptor.TransformFinalBlock(Message, 0, CorArray.Length(Message))

    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub CreateEncryptor_WhenModeIsOFB_ThrowsCryptographicException()
    Dim csp As New TripleDESCryptoServiceProvider
    csp.Mode = CipherMode.OFB

    On Error Resume Next
    csp.CreateEncryptor

    AssertCryptographicException Err
End Sub

Public Sub CreateEncryptor_WhenModeIsCFBAndFeedbackSizeIsNot8_ThrowsCryptographicException()
    Dim csp As New TripleDESCryptoServiceProvider
    csp.Mode = CipherMode.CFB
    csp.FeedbackSize = 16

    On Error Resume Next
    csp.CreateEncryptor

    AssertCryptographicException Err
End Sub

Public Sub TransformBlock_WithInputOffsetLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim csp As New TripleDESCryptoServiceProvider
    Dim Encryptor As ICryptoTransform
    Set Encryptor = csp.CreateEncryptor
    Dim Message(5 To 12) As Byte

    On Error Resume Next
    Encryptor.TransformBlock Message, 4, 8, NewBytes(), 0

    AssertArgumentOutOfRangeException Err, "InputOffset"
End Sub

Public Sub TransformBlock_WithInputCountLessThanZero_ThrowsArgumentException()
    Dim csp As New TripleDESCryptoServiceProvider
    Dim Encryptor As ICryptoTransform
    Set Encryptor = csp.CreateEncryptor
    Dim Message(5 To 12) As Byte

    On Error Resume Next
    Encryptor.TransformBlock Message, 4, -1, NewBytes(), 0

    AssertArgumentException Err, "InputCount"
End Sub
