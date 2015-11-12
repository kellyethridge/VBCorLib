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
    Test.Use CipherMode.ECB, "26 1E 57 8E C9 62 BF B8 3E 96", "00 00 00 00 00 00 00 00", 80, "F9 9A 3A DB 00 3B 7A EB 81 E3 6B A9 E5 37 10 D1 F9 9A 3A DB 00 3B 7A EB 81 E3 6B A9 E5 37 10 D1"
    Test.Use CipherMode.CBC, "26 1E 57 8E C9 62 BF B8 3E 96", "01 02 03 04 05 06 07 08", 40, "71 2D 11 99 C9 A0 78 4F CD F1 1E 3D FD 21 7E DB B2 6E 0D A4 72 BC 31 51 48 EF 4E 68 3B DC CD 7D"
End Sub

Public Sub TransformFinalBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByVal EffectiveKeySize As Long, ByRef Expected As String)
    Dim csp         As New RC2CryptoServiceProvider
    Dim Encryptor   As ICryptoTransform
    Dim Message()   As Byte
    csp.Mode = CipherMode
    csp.Padding = PaddingMode.None
    csp.EffectiveKeySize = EffectiveKeySize
    Set Encryptor = csp.CreateEncryptor(HexToBytes(Key), HexToBytes(IV))
    Message = HexToBytes("11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00 11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00")

    Dim Actual() As Byte
    Actual = Encryptor.TransformFinalBlock(Message, 0, CorArray.Length(Message))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub TransformBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByVal EffectiveKeySize As Long, ByRef Expected As String)
    Dim csp         As New RC2CryptoServiceProvider
    Dim Encryptor   As ICryptoTransform
    Dim Message()   As Byte
    csp.Mode = CipherMode
    csp.Padding = PaddingMode.None
    csp.EffectiveKeySize = EffectiveKeySize
    Set Encryptor = csp.CreateEncryptor(HexToBytes(Key), HexToBytes(IV))
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
    Dim csp As New RC2CryptoServiceProvider
    csp.Mode = CipherMode.OFB
    
    On Error Resume Next
    csp.CreateEncryptor
    
    AssertCryptographicException Err
End Sub

Public Sub CreateEncryptor_WhenModeIsCFBAndFeedbackSizeIsNot8_ThrowsCryptographicException()
    Dim csp As New RC2CryptoServiceProvider
    csp.Mode = CipherMode.CFB
    csp.FeedbackSize = 16
    
    On Error Resume Next
    csp.CreateEncryptor
    
    AssertCryptographicException Err
End Sub

Public Sub CreateEncryptor_WithECBModeAndKeyOnly_DoesNotThrow()
    Dim csp As New RC2CryptoServiceProvider
    csp.Mode = CipherMode.ECB
    
    On Error Resume Next
    csp.CreateEncryptor csp.Key
        
    AssertNoException Err
End Sub

Public Sub CreateEncryptor_WithNonECBModeAndKeyOnly_ThrowsArgumentException()
    Dim csp As New RC2CryptoServiceProvider
    csp.Mode = CipherMode.CBC
    
    On Error Resume Next
    csp.CreateEncryptor csp.Key
    
    AssertArgumentException Err, "RgbIV"
End Sub

Public Sub CreateEncryptor_WithIVOnly_ThrowsArgumentException()
    Dim csp As New RC2CryptoServiceProvider
    
    On Error Resume Next
    csp.CreateEncryptor RgbIV:=NewBytes()
    
    AssertArgumentException Err, "RgbKey"
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
    Dim csp As New RC2CryptoServiceProvider
    Dim Encryptor As ICryptoTransform
    Set Encryptor = csp.CreateEncryptor
    Dim Message(5 To 12) As Byte
    
    On Error Resume Next
    Encryptor.TransformBlock Message, 4, -1, NewBytes(), 0
    
    AssertArgumentException Err, "InputCount"
End Sub
