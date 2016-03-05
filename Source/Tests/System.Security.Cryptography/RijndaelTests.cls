VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RijndaelTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' RijndaelTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)

End Sub

Public Sub Constructor_WhenCalled_SetsDefaults()
    Dim r As Rijndael
    Set r = Rijndael.Create
    
    Assert.That r.BlockSize, Iz.EqualTo(128), "BlockSize"
    Assert.That r.FeedbackSize, Iz.EqualTo(128), "FeedbackSize"
    Assert.That r.KeySize, Iz.EqualTo(256), "KeySize"
    Assert.That r.Mode, Iz.EqualTo(CipherMode.CBC), "Mode"
    Assert.That r.Padding, Iz.EqualTo(PaddingMode.PKCS7), "Padding"
    Assert.That CorArray.Length(r.LegalBlockSizes), Iz.EqualTo(1), "LegalBlockSizes"
    
    Dim LegalBlockSizes() As KeySizes
    LegalBlockSizes = r.LegalBlockSizes
    Assert.That LegalBlockSizes(0).MaxSize, Iz.EqualTo(256), "LegalBlockSizes.MaxSize"
    Assert.That LegalBlockSizes(0).MinSize, Iz.EqualTo(128), "LegalBlockSizes.MinSize"
    Assert.That LegalBlockSizes(0).SkipSize, Iz.EqualTo(64), "LegalBlockSizes.SkipSize"
    Assert.That CorArray.Length(r.LegalKeySizes), Iz.EqualTo(1), "LegalKeySizes"
    
    Dim LegalKeySizes() As KeySizes
    LegalKeySizes = r.LegalKeySizes
    Assert.That LegalKeySizes(0).MaxSize, Iz.EqualTo(256), "LegalKeySizes.MaxSize"
    Assert.That LegalKeySizes(0).MinSize, Iz.EqualTo(128), "LegalKeySizes.MinSize"
    Assert.That LegalKeySizes(0).SkipSize, Iz.EqualTo(64), "LegalKeySizes.SkipSize"
End Sub

' FIPS197 B
Public Sub Test_FIPS197_AppendixB()
    Dim Key()       As Byte: Key = NewBytes(&H2B, &H7E, &H15, &H16, &H28, &HAE, &HD2, &HA6, &HAB, &HF7, &H15, &H88, &H9, &HCF, &H4F, &H3C)
    Dim IV(0 To 15) As Byte
    Dim Data()      As Byte: Data = NewBytes(&H32, &H43, &HF6, &HA8, &H88, &H5A, &H30, &H8D, &H31, &H31, &H98, &HA2, &HE0, &H37, &H7, &H34)
    Dim Expected()  As Byte: Expected = NewBytes(&H39, &H25, &H84, &H1D, &H2, &HDC, &H9, &HFB, &HDC, &H11, &H85, &H97, &H19, &H6A, &HB, &H32)
    
    Dim r As Rijndael
    Set r = Rijndael.Create
    r.Mode = CipherMode.ECB
    r.KeySize = 128
    r.Padding = PaddingMode.Zeros
    
    Dim Output() As Byte
    ReDim Output(0 To UBound(Data))
    Dim Encryptor As ICryptoTransform
    Set Encryptor = r.CreateEncryptor(Key, IV)
    Encryptor.TransformBlock Data, 0, CorArray.Length(Data), Output, 0
    Assert.That Output, Iz.EqualTo(Expected), "FIPS197 B Encrypt"
    
    Dim Original() As Byte
    ReDim Original(0 To UBound(Output))
    Dim Decryptor As ICryptoTransform
    Set Decryptor = r.CreateDecryptor(Key, IV)
    Decryptor.TransformBlock Output, 0, CorArray.Length(Output), Original, 0
    Assert.That Original, Iz.EqualTo(Data), "FIPS197 B Decrypt"
End Sub

' FIPS197 C.1 AES-128 (Nk=4, Nr=10)
Public Sub Test_FIPS197_AppendixC1()
    Dim Key()       As Byte: Key = NewBytes(&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF)
    Dim IV(0 To 15) As Byte
    Dim Data()      As Byte: Data = NewBytes(&H0, &H11, &H22, &H33, &H44, &H55, &H66, &H77, &H88, &H99, &HAA, &HBB, &HCC, &HDD, &HEE, &HFF)
    Dim Expected()  As Byte: Expected = NewBytes(&H69, &HC4, &HE0, &HD8, &H6A, &H7B, &H4, &H30, &HD8, &HCD, &HB7, &H80, &H70, &HB4, &HC5, &H5A)
    
    Dim r As Rijndael
    Set r = Rijndael.Create
    r.Mode = CipherMode.ECB
    r.KeySize = 128
    r.Padding = PaddingMode.Zeros
    
    Dim Output() As Byte
    ReDim Output(0 To UBound(Data))
    Dim Encryptor As ICryptoTransform
    Set Encryptor = r.CreateEncryptor(Key, IV)
    Encryptor.TransformBlock Data, 0, CorArray.Length(Data), Output, 0
    Assert.That Output, Iz.EqualTo(Expected), "FIPS197 C1 Encrypt"
    
    Dim Original() As Byte
    ReDim Original(0 To UBound(Output))
    Dim Decryptor As ICryptoTransform
    Set Decryptor = r.CreateDecryptor(Key, IV)
    Decryptor.TransformBlock Output, 0, CorArray.Length(Output), Original, 0
    Assert.That Original, Iz.EqualTo(Data), "FIPS197 C1 Decrypt"
End Sub

' FIPS197 C.2 AES-192 (Nk=6, Nr=12)
Public Sub Test_FIPS197_AppendixC2()
    Dim Key()       As Byte: Key = NewBytes(&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17)
    Dim IV(0 To 15) As Byte
    Dim Data()      As Byte: Data = NewBytes(&H0, &H11, &H22, &H33, &H44, &H55, &H66, &H77, &H88, &H99, &HAA, &HBB, &HCC, &HDD, &HEE, &HFF)
    Dim Expected()  As Byte: Expected = NewBytes(&HDD, &HA9, &H7C, &HA4, &H86, &H4C, &HDF, &HE0, &H6E, &HAF, &H70, &HA0, &HEC, &HD, &H71, &H91)
    
    Dim r As Rijndael
    Set r = Rijndael.Create
    r.Mode = CipherMode.ECB
    r.KeySize = 192
    r.Padding = PaddingMode.Zeros
    
    Dim Output() As Byte
    ReDim Output(0 To UBound(Data))
    Dim Encryptor As ICryptoTransform
    Set Encryptor = r.CreateEncryptor(Key, IV)
    Encryptor.TransformBlock Data, 0, CorArray.Length(Data), Output, 0
    Assert.That Output, Iz.EqualTo(Expected), "FIPS197 C2 Encrypt"
    
    Dim Original() As Byte
    ReDim Original(0 To UBound(Output))
    Dim Decryptor As ICryptoTransform
    Set Decryptor = r.CreateDecryptor(Key, IV)
    Decryptor.TransformBlock Output, 0, CorArray.Length(Output), Original, 0
    Assert.That Original, Iz.EqualTo(Data), "FIPS197 C2 Decrypt"
End Sub

' C.3 AES-256 (Nk=8, Nr=14)
Public Sub Test_FIPS197_AppendixC3()
    Dim Key()       As Byte: Key = NewBytes(&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F)
    Dim IV(0 To 15) As Byte
    Dim Data()      As Byte: Data = NewBytes(&H0, &H11, &H22, &H33, &H44, &H55, &H66, &H77, &H88, &H99, &HAA, &HBB, &HCC, &HDD, &HEE, &HFF)
    Dim Expected()  As Byte: Expected = NewBytes(&H8E, &HA2, &HB7, &HCA, &H51, &H67, &H45, &HBF, &HEA, &HFC, &H49, &H90, &H4B, &H49, &H60, &H89)
    
    Dim r As Rijndael
    Set r = Rijndael.Create
    r.Mode = CipherMode.ECB
    r.KeySize = 256
    r.Padding = PaddingMode.Zeros
    
    Dim Output() As Byte
    ReDim Output(0 To UBound(Data))
    Dim Encryptor As ICryptoTransform
    Set Encryptor = r.CreateEncryptor(Key, IV)
    Encryptor.TransformBlock Data, 0, CorArray.Length(Data), Output, 0
    Assert.That Output, Iz.EqualTo(Expected), "FIPS197 C3 Encrypt"
    
    Dim Original() As Byte
    ReDim Original(0 To UBound(Output))
    Dim Decryptor As ICryptoTransform
    Set Decryptor = r.CreateDecryptor(Key, IV)
    Decryptor.TransformBlock Output, 0, CorArray.Length(Output), Original, 0
    Assert.That Original, Iz.EqualTo(Data), "FIPS197 C3 Decrypt"
End Sub

Public Sub Changing_KeySize()
    Dim OriginalIV() As Byte
    Dim r As Rijndael
    Set r = Rijndael.Create
    OriginalIV = r.IV
    
    Dim LegalSizes As KeySizes
    Set LegalSizes = r.LegalKeySizes()(0)
    Dim KeySize As Long
    For KeySize = LegalSizes.MinSize To LegalSizes.MaxSize Step LegalSizes.SkipSize
        r.KeySize = KeySize
        Assert.That KeySize \ 8, Iz.EqualTo(CorArray.Length(r.Key))
        Assert.That r.IV, Iz.EqualTo(OriginalIV)
    Next
End Sub

Public Sub Changing_BlockSize()
    Dim OriginalKey() As Byte
    Dim r As Rijndael
    Set r = Rijndael.Create
    OriginalKey = r.Key
    
    Dim LegalSizes As KeySizes
    Set LegalSizes = r.LegalBlockSizes()(0)
    Dim BlockSize As Long
    For BlockSize = LegalSizes.MinSize To LegalSizes.MaxSize Step LegalSizes.SkipSize
        r.BlockSize = BlockSize
        Assert.That r.Key, Iz.EqualTo(OriginalKey)
        Assert.That BlockSize \ 8, Iz.EqualTo(CorArray.Length(r.IV))
    Next
End Sub

Public Sub FeedbackSize_WhenSetNotMultipleOf8_ThrowsCryptographicException()
    Dim r As Rijndael
    Set r = Rijndael.Create
    
    On Error Resume Next
    r.FeedbackSize = 9
    
    AssertCryptographicException Err
End Sub













