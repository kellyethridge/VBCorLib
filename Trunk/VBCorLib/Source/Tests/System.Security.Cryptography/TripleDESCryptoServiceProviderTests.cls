VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TripleDESCryptoServiceProviderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TripleDESCryptoServiceProviderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "IsWeakKey_WithValue_ReturnsExpected"
            Test.Use "0000000000000000ffffffffffffffff", False
            Test.Use "ffffffffffffffffffffffffffffffff", True
            Test.Use "0000000000000000ffffffffffffffff0000000000000000", False
            Test.Use "ffffffffffffffffffffffffffffffff0000000000000000", True
            Test.Use "0000000000000000ffffffffffffffffffffffffffffffff", True
            
    End Select
End Sub

Public Sub Constructor_WhenCalled_SetsDefaults()
    Dim csp As New TripleDESCryptoServiceProvider
    
    Assert.That csp.KeySize, Iz.EqualTo(192)
    Assert.That csp.BlockSize, Iz.EqualTo(64)
End Sub

Public Sub Constructor_WithOutKeyAndIV_AutoGenerates()
    Dim csp As New TripleDESCryptoServiceProvider
    
    Assert.That csp.Key, Iz.Not.Empty
    Assert.That csp.IV, Iz.Not.Empty
    Assert.That CorArray.Length(csp.Key), Iz.EqualTo(24)
    Assert.That CorArray.Length(csp.IV), Iz.EqualTo(8)
End Sub

Public Sub Key_WhenSetToNull_ThrowsArgumentNullException()
    Dim csp As New TripleDESCryptoServiceProvider
    Dim NullBytes() As Byte
    
    On Error Resume Next
    csp.Key = NullBytes
    
    AssertArgumentNullException Err, "Key"
End Sub

Public Sub Key_WithNonLegalSize_ThrowsArgumentException()
    Dim csp As New TripleDESCryptoServiceProvider
    Dim Key() As Byte
    ReDim Key(7)
    
    On Error Resume Next
    csp.Key = Key
    
    AssertArgumentException Err, "Key"
End Sub

Public Sub Key_WhenSetTo128Bits_UpdatesKey()
    Dim csp As New TripleDESCryptoServiceProvider
    Dim Key() As Byte
    Key = NewBytes(1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 1, 2, 3, 4, 5, 6)
    
    csp.Key = Key
    
    Assert.That csp.Key, Iz.EqualTo(Key)
    Assert.That csp.KeySize, Iz.EqualTo(128)
End Sub

Public Sub KeySize_WhenSetTo128Bits_UpdatesKey()
    Dim csp As New TripleDESCryptoServiceProvider
    
    csp.KeySize = 128
    
    Assert.That CorArray.Length(csp.Key), Iz.EqualTo(16)
End Sub

Public Sub IV_WhenSet_StoresIV()
    Dim csp As New TripleDESCryptoServiceProvider
    Dim IV() As Byte
    IV = NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    csp.IV = IV
    
    Assert.That csp.IV, Iz.EqualTo(IV)
End Sub

Public Sub IsWeakKey_WithValue_ReturnsExpected(ByRef Key As String, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = TripleDES.IsWeakKey(HexToBytes(Key))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsWeakKey_WithInvalidKeySize_ThrowsCryptographicException()
    On Error Resume Next
    
    TripleDES.IsWeakKey NewBytes(0, 0, 0, 0, 0, 0, 0, 0)
    
    AssertCryptographicException Err
End Sub


