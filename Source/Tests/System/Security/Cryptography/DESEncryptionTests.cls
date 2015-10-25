VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DESEncryptionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   DESEncryptionTests
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
            Test.Use PaddingMode.None, "05c9c4cafb9937d95bbfbedfc5d77f19a6cd5a5dab188a33dfd8979fc4b7b2be"
            Test.Use PaddingMode.ANSIX923, "05c9c4cafb9937d95bbfbedfc5d77f19a6cd5a5dab188a33dfd8979fc4b7b2be030116f7e552e7b6"
            Test.Use PaddingMode.PKCS7, "05c9c4cafb9937d95bbfbedfc5d77f19a6cd5a5dab188a33dfd8979fc4b7b2befeb959b7d4642fcb"
            
    End Select
End Sub

Private Sub AddTests(ByVal Test As TestCaseCollector)
    Test.Use CipherMode.ECB, "12345678", "00000000", "05c9c4cafb9937d95bbfbedfc5d77f19a6cd5a5dab188a33dfd8979fc4b7b2be"
    Test.Use CipherMode.CBC, "12345678", "abcdefgh", "6ca9470c849d1cc1a59ffc148f1cb5e9cf1f5c0328a7e8756387ff4d0fe46050"
End Sub

Public Sub TransformFinalBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByRef Expected As String)
    Dim csp         As New DESCryptoServiceProvider
    Dim Encryptor   As ICryptoTransform
    Dim Message()   As Byte
    csp.Mode = CipherMode
    csp.Padding = PaddingMode.None
    Set Encryptor = csp.CreateEncryptor(Encoding.ASCII.GetBytes(Key), Encoding.ASCII.GetBytes(IV))
    Message = Encoding.ASCII.GetBytes("This is the message to encrypt!!")

    Dim Actual() As Byte
    Actual = Encryptor.TransformFinalBlock(Message, 0, CorArray.Length(Message))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub TransformBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByRef Expected As String)
    Dim csp         As New DESCryptoServiceProvider
    Dim Encryptor   As ICryptoTransform
    Dim Message()   As Byte
    csp.Mode = CipherMode
    csp.Padding = PaddingMode.None
    Set Encryptor = csp.CreateEncryptor(Encoding.ASCII.GetBytes(Key), Encoding.ASCII.GetBytes(IV))
    Message = Encoding.ASCII.GetBytes("This is the message to encrypt!!")
    
    Dim Actual          As Long
    Dim ActualOutput()  As Byte
    ReDim ActualOutput(UBound(Message))
    Actual = Encryptor.TransformBlock(Message, 0, CorArray.Length(Message), ActualOutput, 0)
    
    Assert.That Actual, Iz.EqualTo(CorArray.Length(Message))
    Assert.That ActualOutput, Iz.EqualTo(HexToBytes(Expected))
    Assert.That Encryptor.TransformFinalBlock(NewBytes(), 0, 0), Iz.Empty
End Sub

Public Sub TransformFinalBlock_WithPaddingMode_ReturnsExpected(ByVal Mode As PaddingMode, ByRef Expected As String)
    Dim csp As New DESCryptoServiceProvider
    Dim Encryptor As ICryptoTransform
    Dim Message()   As Byte
    csp.Mode = CipherMode.ECB
    csp.Padding = Mode
    Set Encryptor = csp.CreateEncryptor(Encoding.ASCII.GetBytes("12345678"), Encoding.ASCII.GetBytes("00000000"))
    Message = Encoding.ASCII.GetBytes("This is the message to encrypt!!")

    Dim Actual() As Byte
    Actual = Encryptor.TransformFinalBlock(Message, 0, CorArray.Length(Message))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub CreateEncryptor_WhenModeIsOFB_ThrowsCryptographicException()
    Dim csp As New DESCryptoServiceProvider
    csp.Mode = CipherMode.OFB
    
    On Error Resume Next
    csp.CreateEncryptor
    
    AssertCryptographicException Err
End Sub

Public Sub CreateEncryptor_WhenModeIsCFBAndFeedbackSizeIsNot8_ThrowsCryptographicException()
    Dim csp As New DESCryptoServiceProvider
    csp.Mode = CipherMode.CFB
    csp.FeedbackSize = 16
    
    On Error Resume Next
    csp.CreateEncryptor
    
    AssertCryptographicException Err
End Sub

Public Sub TransformBlock_WithInputOffsetLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim csp As New DESCryptoServiceProvider
    Dim Encryptor As ICryptoTransform
    Set Encryptor = csp.CreateEncryptor
    Dim Message(5 To 12) As Byte
    
    On Error Resume Next
    Encryptor.TransformBlock Message, 4, 8, NewBytes(), 0
    
    AssertArgumentOutOfRangeException Err, "InputOffset"
End Sub

Public Sub TransformBlock_WithInputCountLessThanZero_ThrowsArgumentException()
    Dim csp As New DESCryptoServiceProvider
    Dim Encryptor As ICryptoTransform
    Set Encryptor = csp.CreateEncryptor
    Dim Message(5 To 12) As Byte
    
    On Error Resume Next
    Encryptor.TransformBlock Message, 4, -1, NewBytes(), 0
    
    AssertArgumentException Err, "InputCount"
End Sub


























