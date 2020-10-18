VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RC2DecryptionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' RC2DecryptionTests
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

Public Sub TransformFinalBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByRef Crypted As String)
    Dim csp         As New RC2CryptoServiceProvider
    Dim Decryptor   As ICryptoTransform
    Dim MCrypt()    As Byte
    csp.Mode = CipherMode
    csp.Padding = PaddingMode.None
    Set Decryptor = csp.CreateDecryptor(HexToBytes(Key), HexToBytes(IV))
    MCrypt = HexToBytes(Crypted)
    
    Dim Actual() As Byte
    Actual = Decryptor.TransformFinalBlock(MCrypt, 0, CorArray.Length(MCrypt))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes("11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00 11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00"))
End Sub

Public Sub TransformBlock_WithCipherMode_ReturnsExpected(ByVal CipherMode As CipherMode, ByRef Key As String, ByRef IV As String, ByRef Crypted As String)
    Dim csp         As New RC2CryptoServiceProvider
    Dim Decryptor   As ICryptoTransform
    Dim MCrypt()    As Byte
    csp.Mode = CipherMode
    csp.Padding = PaddingMode.None
    Set Decryptor = csp.CreateDecryptor(HexToBytes(Key), HexToBytes(IV))
    MCrypt = HexToBytes(Crypted)
    
    Dim Actual          As Long
    Dim ActualOutput()  As Byte
    ReDim ActualOutput(UBound(MCrypt))
    Actual = Decryptor.TransformBlock(MCrypt, 0, CorArray.Length(MCrypt), ActualOutput, 0)
    
    Assert.That Actual, Iz.EqualTo(CorArray.Length(MCrypt))
    Assert.That ActualOutput, Iz.EqualTo(HexToBytes("11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00 11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00"))
    Assert.That Decryptor.TransformFinalBlock(NewBytes(), 0, 0), Iz.Empty
End Sub

