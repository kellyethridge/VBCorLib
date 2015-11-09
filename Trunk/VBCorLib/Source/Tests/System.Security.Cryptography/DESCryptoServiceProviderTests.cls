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
            
        Case "BlockSize_WithInvalidValue_ThrowsCryptographicException"
            Test.Use 32
            Test.Use 96
            
        Case "FeedbackSize_WithInvalidValue_ThrowsCryptographiException"
            Test.Use 0
            Test.Use 72
            Test.Use 63
            
        Case "FeedbackSize_WithValidValue_DoesNotThrow"
            Test.Use 8
            Test.Use 32
            Test.Use 64
            
    End Select
End Sub

Public Sub Constructor_WithOutKeyAndIV_AutoGenerates()
    Dim csp As New DESCryptoServiceProvider
    
    Assert.That csp.Key, Iz.Not.Empty
    Assert.That csp.IV, Iz.Not.Empty
    Assert.That CorArray.Length(csp.Key), Iz.EqualTo(8)
    Assert.That CorArray.Length(csp.IV), Iz.EqualTo(8)
End Sub

Public Sub BlockSize_WithInvalidValue_ThrowsCryptographicException(ByVal InvalidValue As Long)
    Dim csp As New DESCryptoServiceProvider
    
    On Error Resume Next
    csp.BlockSize = InvalidValue
    
    AssertCryptographicException Err
End Sub

Public Sub FeedbackSize_WithInvalidValue_ThrowsCryptographiException(ByVal InvalidValue As Long)
    Dim csp As New DESCryptoServiceProvider
    
    On Error Resume Next
    csp.FeedbackSize = InvalidValue
    
    AssertCryptographicException Err
End Sub

Public Sub FeedbackSize_WithValidValue_DoesNotThrow(ByVal ValidValue As Long)
    Dim csp As New DESCryptoServiceProvider
    
    csp.FeedbackSize = ValidValue
End Sub

Public Sub Key_WhenSetToNull_ThrowsArgumentNullException()
    Dim csp As New DESCryptoServiceProvider
    Dim NullBytes() As Byte
    
    On Error Resume Next
    csp.Key = NullBytes
    
    AssertArgumentNullException Err, "Key"
End Sub

Public Sub Key_WhenSetToInvalidSize_ThrowsArgumentException()
    Dim csp As New DESCryptoServiceProvider
    Dim Key() As Byte
    ReDim Key(0)
    
    On Error Resume Next
    csp.Key = Key
    
    AssertArgumentException Err, "Key"
End Sub

Public Sub Key_WhenSet_StoresKey()
    Dim csp As New DESCryptoServiceProvider
    Dim Key() As Byte
    Key = NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    csp.Key = Key
    
    Assert.That csp.Key, Iz.EqualTo(Key)
End Sub

Public Sub IV_WhenSet_StoresIV()
    Dim csp As New DESCryptoServiceProvider
    Dim IV() As Byte
    IV = NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    csp.IV = IV
    
    Assert.That csp.IV, Iz.EqualTo(IV)
End Sub

Public Sub KeySize_WithNonLegalSize_ThrowsCrypgraphicException()
    Dim csp As New DESCryptoServiceProvider
    
    On Error Resume Next
    csp.KeySize = 128
    
    AssertCryptographicException Err
End Sub

Public Sub IV_WhenSetToInvalidSize_ThrowsCryptographicException()
    Dim csp As New DESCryptoServiceProvider
    Dim NotBlockSizeArray() As Byte
    ReDim NotBlockSizeArray(1 To 7)
    
    On Error Resume Next
    csp.IV = NotBlockSizeArray
    
    AssertCryptographicException Err
End Sub

Public Sub GenerateIV_WhenCalled_GeneratesCorrectSizeArray()
    Dim csp As New DESCryptoServiceProvider
    
    csp.GenerateIV
    
    Assert.That CorArray.Length(csp.IV), Iz.EqualTo(8)
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

