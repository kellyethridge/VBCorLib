VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RC2EncryptionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' RC2EncryptionTests
'
' Test vectors from http://cryptomanager.com/tv.html#ciphs
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
        
    End Select
End Sub

Private Sub AddTests(ByVal Test As TestCaseCollector)
    Test.Use CipherMode.ECB, "26 1E 57 8E C9 62 BF B8 3E 96", "00 00 00 00 00 00 00 00", "F9 9A 3A DB 00 3B 7A EB 81 E3 6B A9 E5 37 10 D1 F9 9A 3A DB 00 3B 7A EB 81 E3 6B A9 E5 37 10 D1"
End Sub

Public Sub TransformFinalBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByRef Expected As String)
    Dim Csp         As New RC2CryptoServiceProvider
    Dim Encryptor   As ICryptoTransform
    Dim Message()   As Byte
    Csp.Mode = CipherMode
    Csp.Padding = PaddingMode.None
    Set Encryptor = Csp.CreateEncryptor(HexToBytes(Key), HexToBytes(IV))
    Message = HexToBytes("11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00 11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00")

    Dim Actual() As Byte
    Actual = Encryptor.TransformFinalBlock(Message, 0, CorArray.Length(Message))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub TransformBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByRef Expected As String)
    Dim Csp         As New RC2CryptoServiceProvider
    Dim Encryptor   As ICryptoTransform
    Dim Message()   As Byte
    Csp.Mode = CipherMode
    Csp.Padding = PaddingMode.None
    Set Encryptor = Csp.CreateEncryptor(HexToBytes(Key), HexToBytes(IV))
    Message = HexToBytes("11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00 11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00")
    
    Dim Actual          As Long
    Dim ActualOutput()  As Byte
    ReDim ActualOutput(UBound(Message))
    Actual = Encryptor.TransformBlock(Message, 0, CorArray.Length(Message), ActualOutput, 0)
    
    Assert.That Actual, Iz.EqualTo(CorArray.Length(Message))
    Assert.That ActualOutput, Iz.EqualTo(HexToBytes(Expected))
    Assert.That Encryptor.TransformFinalBlock(NewBytes(), 0, 0), Iz.Empty
End Sub

Public Sub CreateEncryptor_WhenModeIsOFB_ThrowsCryptographicException()
    Dim Csp As New RC2CryptoServiceProvider
    Csp.Mode = CipherMode.OFB
    
    On Error Resume Next
    Csp.CreateEncryptor
    
    AssertCryptographicException Err
End Sub

Public Sub CreateEncryptor_WhenModeIsCFBAndFeedbackSizeIsNot8_ThrowsCryptographicException()
    Dim Csp As New RC2CryptoServiceProvider
    Csp.Mode = CipherMode.CFB
    Csp.FeedbackSize = 16
    
    On Error Resume Next
    Csp.CreateEncryptor
    
    AssertCryptographicException Err
End Sub

Public Sub CreateEncryptor_WithECBModeAndKeyOnly_DoesNotThrow()
    Dim Csp As New RC2CryptoServiceProvider
    Csp.Mode = CipherMode.ECB
    
    On Error Resume Next
    Csp.CreateEncryptor Csp.Key
        
    AssertNoException Err
End Sub

Public Sub CreateEncryptor_WithNonECBModeAndKeyOnly_ThrowsArgumentException()
    Dim Csp As New RC2CryptoServiceProvider
    Csp.Mode = CipherMode.CBC
    
    On Error Resume Next
    Csp.CreateEncryptor Csp.Key
    
    AssertArgumentException Err, "RgbIV"
End Sub

Public Sub CreateEncryptor_WithIVOnly_ThrowsArgumentException()
    Dim Csp As New RC2CryptoServiceProvider
    
    On Error Resume Next
    Csp.CreateEncryptor RgbIV:=NewBytes()
    
    AssertArgumentException Err, "RgbKey"
End Sub

Public Sub TransformBlock_WithInputOffsetLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Csp As New DESCryptoServiceProvider
    Dim Encryptor As ICryptoTransform
    Set Encryptor = Csp.CreateEncryptor
    Dim Message(5 To 12) As Byte
    
    On Error Resume Next
    Encryptor.TransformBlock Message, 4, 8, NewBytes(), 0
    
    AssertArgumentOutOfRangeException Err, "InputOffset"
End Sub

Public Sub TransformBlock_WithInputCountLessThanZero_ThrowsArgumentException()
    Dim Csp As New RC2CryptoServiceProvider
    Dim Encryptor As ICryptoTransform
    Set Encryptor = Csp.CreateEncryptor
    Dim Message(5 To 12) As Byte
    
    On Error Resume Next
    Encryptor.TransformBlock Message, 5, -1, NewBytes(), 0
    
    AssertArgumentException Err, "InputCount"
End Sub
