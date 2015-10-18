VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DESCryptoServiceProviderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' DESCryptoServiceProviderTests
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
            Test.Use NewBytes(1, 2, 3, 4, 5, 6, 7, 8), False
            Test.Use NewBytes(2, 2, 2, 2, 0, 0, 0, 0), False
            Test.Use NewBytes(0, 0, 0, 0, 0, 0, 0, 0), True
            Test.Use NewBytes(&HFE, &HFE, &HFE, &HFE, &HFE, &HFE, &HFE, &HFE), True
            Test.Use NewBytes(&H1F, &H1F, &H1F, &H1F, &HE, &HE, &HE, &HE), True
            Test.Use NewBytes(&HE0, &HE0, &HE0, &HE0, &HF1, &HF1, &HF1, &HF1), True
        
        Case "IsSemiWeakKey_WithValue_ReturnsExpected"
            Test.Use NewBytes(123, 0, 0, 0, 0, 0, 0, 0), False
            Test.Use HexToBytes("E001E001F101F101"), True
            Test.Use HexToBytes("01E001E001F101F1"), True
            Test.Use HexToBytes("FE1FFE1FFE0EFE0E"), True
            Test.Use HexToBytes("1FFE1FFE0EFE0EFE"), True
            Test.Use HexToBytes("E01FE01FF10EF10E"), True
            Test.Use HexToBytes("1FE01FE00EF10EF1"), True
            Test.Use HexToBytes("01FE01FE01FE01FE"), True
            Test.Use HexToBytes("FE01FE01FE01FE01"), True
            Test.Use HexToBytes("011F011F010E010E"), True
            Test.Use HexToBytes("1F011F010E010E01"), True
            Test.Use HexToBytes("E0FEE0FEF1FEF1FE"), True
            Test.Use HexToBytes("FEE0FEE0FEF1FEF1"), True
            
    End Select
End Sub

Public Sub Constructor_WithOutKeyAndIV_AutoGenerates()
    Dim Csp As New DESCryptoServiceProvider
    
    Assert.That Csp.Key, Iz.Not.Empty
    Assert.That Csp.IV, Iz.Not.Empty
    Assert.That CorArray.Length(Csp.Key), Iz.EqualTo(8)
    Assert.That CorArray.Length(Csp.IV), Iz.EqualTo(8)
End Sub

Public Sub Key_WhenSetToNull_ThrowsArgumentNullException()
    Dim Csp As New DESCryptoServiceProvider
    Dim NullBytes() As Byte
    
    On Error Resume Next
    Csp.Key = NullBytes
    
    AssertArgumentNullException Err, "Key"
End Sub

Public Sub Key_WhenSetToInvalidSize_ThrowsArgumentException()
    Dim Csp As New DESCryptoServiceProvider
    Dim Key() As Byte
    ReDim Key(0)
    
    On Error Resume Next
    Csp.Key = Key
    
    AssertArgumentException Err, "Key"
End Sub

Public Sub Key_WhenSet_StoresKey()
    Dim Csp As New DESCryptoServiceProvider
    Dim Key() As Byte
    Key = NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Csp.Key = Key
    
    Assert.That Csp.Key, Iz.EqualTo(Key)
End Sub

Public Sub IV_WhenSet_StoresIV()
    Dim Csp As New DESCryptoServiceProvider
    Dim IV() As Byte
    IV = NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Csp.IV = IV
    
    Assert.That Csp.IV, Iz.EqualTo(IV)
End Sub

Public Sub KeySize_WithNonLegalSize_ThrowsCrypgraphicException()
    Dim Csp As New DESCryptoServiceProvider
    
    On Error Resume Next
    Csp.KeySize = 128
    
    AssertCryptographicException Err
End Sub

Public Sub IV_WhenSetToInvalidSize_ThrowsCryptographicException()
    Dim Csp As New DESCryptoServiceProvider
    
    On Error Resume Next
    Csp.IV = NewBytes(1, 2, 3)
    
    AssertCryptographicException Err
End Sub

Public Sub IsWeakKey_WithValue_ReturnsExpected(ByRef Key() As Byte, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = DES.IsWeakKey(Key)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsSemiWeakKey_WithValue_ReturnsExpected(ByRef Key() As Byte, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = DES.IsSemiWeakKey(Key)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

