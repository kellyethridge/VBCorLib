VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RijndaelManagedTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' RijndaelManagedTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Encrypt_WithECBMode_ReturnsExpected"
            Test.Use 128, NewBytes(&H69, &HC4, &HE0, &HD8, &H6A, &H7B, &H4, &H30, &HD8, &HCD, &HB7, &H80, &H70, &HB4, &HC5, &H5A)
            Test.Use 192, NewBytes(&HDD, &HA9, &H7C, &HA4, &H86, &H4C, &HDF, &HE0, &H6E, &HAF, &H70, &HA0, &HEC, &HD, &H71, &H91)
            Test.Use 256, NewBytes(&H8E, &HA2, &HB7, &HCA, &H51, &H67, &H45, &HBF, &HEA, &HFC, &H49, &H90, &H4B, &H49, &H60, &H89)
            
        Case "Decrypt_WithECBMode_ReturnsExpected"
            Test.Use 128, NewBytes(&H69, &HC4, &HE0, &HD8, &H6A, &H7B, &H4, &H30, &HD8, &HCD, &HB7, &H80, &H70, &HB4, &HC5, &H5A)
            Test.Use 192, NewBytes(&HDD, &HA9, &H7C, &HA4, &H86, &H4C, &HDF, &HE0, &H6E, &HAF, &H70, &HA0, &HEC, &HD, &H71, &H91)
            Test.Use 256, NewBytes(&H8E, &HA2, &HB7, &HCA, &H51, &H67, &H45, &HBF, &HEA, &HFC, &H49, &H90, &H4B, &H49, &H60, &H89)
        
        Case "CreateEncryptor_WithInvalidIVSize_ThrowsCryptographicException"
            Test.Use 0
            Test.Use 7
            Test.Use 512
        
        Case "CreateDecryptor_WithInvalidIVSize_ThrowsCryptographicException"
            Test.Use 0
            Test.Use 7
            Test.Use 512
        
    End Select
End Sub

Public Sub Encrypt_WithECBMode_ReturnsExpected(ByVal KeySize As Long, ByRef Expected() As Byte)
    Dim Transform As ICryptoTransform
    Set Transform = CreateTransform(KeySize, "CreateEncryptor")
    
    Dim ActualCount As Long
    Dim PlainText() As Byte
    Dim CipherText() As Byte
    
    PlainText = CreatePlainText
    ReDim CipherText(0 To UBound(PlainText))
    ActualCount = Transform.TransformBlock(PlainText, 0, CorArray.Length(PlainText), CipherText, 0)
    
    Assert.That ActualCount, Iz.EqualTo(16)
    Assert.That CipherText, Iz.EqualTo(Expected)
End Sub

Public Sub Encrypt_WithNullInputBuffer_ThrowsArgumentNullException()
    Dim Transform As ICryptoTransform
    Dim Encryption As New RijndaelManaged
    Dim Output(0) As Byte
    
    Set Transform = Encryption.CreateEncryptor
    
    On Error Resume Next
    Transform.TransformBlock NullBytes, 0, 1, Output, 0
    
    AssertArgumentNullException Err, "InputBuffer"
End Sub

Public Sub Decrypt_WithECBMode_ReturnsExpected(ByVal KeySize As Long, ByRef CipherText() As Byte)
    Dim Transform As ICryptoTransform
    Set Transform = CreateTransform(KeySize, "CreateDecryptor")
    
    Dim ActualCount As Long
    Dim PlainText(0 To 15) As Byte
    ActualCount = Transform.TransformBlock(CipherText, 0, CorArray.Length(CipherText), PlainText, 0)
    
    Assert.That ActualCount, Iz.EqualTo(16)
    Assert.That PlainText, Iz.EqualTo(CreatePlainText)
End Sub

Public Sub CreateEncryptor_WithNullIV_DoesNotAlterIV()
    Dim r As New RijndaelManaged
    Dim Encryptor As ICryptoTransform
    Dim Decryptor As ICryptoTransform
        
    Set Encryptor = r.CreateEncryptor(r.Key, NullBytes)
    Set Decryptor = r.CreateDecryptor(r.Key, r.IV)
    
    Dim Data() As Byte
    ReDim Data(0 To Encryptor.InputBlockSize - 1)
    Dim EncData() As Byte
    Dim Actual() As Byte
    EncData = Encryptor.TransformFinalBlock(Data, 0, CorArray.Length(Data))
    Actual = Decryptor.TransformFinalBlock(EncData, 0, CorArray.Length(EncData))
    
    Assert.That Actual, Iz.Not.EqualTo(Data)
End Sub

Public Sub CreateEncryptor_WithKeyAndIV_DoesNotAlterOriginal()
    Dim r As New RijndaelManaged
    Dim OriginalKey() As Byte
    Dim OriginalIV() As Byte
    Dim Key() As Byte
    Dim IV() As Byte
    
    OriginalKey = r.Key
    OriginalIV = r.IV
    Key = r.Key
    IV = r.IV
    CorArray.Reverse Key
    CorArray.Reverse IV
    
    Assert.That r.CreateEncryptor(Key, IV), Iz.Not.Nothing
    
    Assert.That r.Key, Iz.EqualTo(OriginalKey)
    Assert.That r.IV, Iz.EqualTo(OriginalIV)
End Sub

Public Sub CreateDecryptor_WithNullIV_DoesNotAlterIV()
    Dim r As New RijndaelManaged
    Dim Encryptor As ICryptoTransform
    Set Encryptor = r.CreateEncryptor(r.Key, r.IV)
    ReDim Data(0 To Encryptor.InputBlockSize - 1) As Byte
    Dim EncData() As Byte
    EncData = Encryptor.TransformFinalBlock(Data, 0, CorArray.Length(Data))
    
    Dim Actual() As Byte
    Dim Decryptor As ICryptoTransform
    Set Decryptor = r.CreateDecryptor(r.Key, NullBytes)
    Actual = Decryptor.TransformFinalBlock(EncData, 0, CorArray.Length(EncData))
    
    Assert.That Actual, Iz.Not.EqualTo(Data)
End Sub

Public Sub CreateDecryptor_WithKeyAndIV_DoesNotAlterOriginal()
    Dim r As New RijndaelManaged
    Dim OriginalKey() As Byte
    Dim OriginalIV() As Byte
    Dim Key() As Byte
    Dim IV() As Byte
    
    OriginalKey = r.Key
    OriginalIV = r.IV
    Key = r.Key
    IV = r.IV
    CorArray.Reverse Key
    CorArray.Reverse IV
    
    Assert.That r.CreateDecryptor(Key, IV), Iz.Not.Nothing
    
    Assert.That r.Key, Iz.EqualTo(OriginalKey)
    Assert.That r.IV, Iz.EqualTo(OriginalIV)
End Sub

Public Sub CreateEncryptor_WithInvalidIVSize_ThrowsCryptographicException(ByVal Size As Long)
    Dim r As New RijndaelManaged
    Dim IV() As Byte
    If Size >= 0 Then
        IV = CorArray.CreateInstance(vbByte, Size)
    End If
    
    On Error Resume Next
    r.CreateEncryptor r.Key, IV
    
    AssertCryptographicException Err
End Sub

Public Sub CreateDecryptor_WithInvalidIVSize_ThrowsCryptographicException(ByVal Size As Long)
    Dim r As New RijndaelManaged
    Dim IV() As Byte
    If Size >= 0 Then
        IV = CorArray.CreateInstance(vbByte, Size)
    End If
    
    On Error Resume Next
    r.CreateDecryptor r.Key, IV
    
    AssertCryptographicException Err
End Sub

Public Sub Encrypt_MonoBug7193()
    Const Size As Long = 23
    Dim Original(0 To Size - 1) As Byte
    Dim Expected() As Byte
    Expected = NewBytes(&HDC, &HA8, &H39, &H5C, &HA1, &H89, &H3B, &H5, &HFA, &HD8, &HB5, &H76, &H5F, &H8F, &H40, &HCF, &HA7, &HFF, &H86, &HE6, &H30, &H67, &H6B)
    Dim EncData() As Byte
    Dim DecData() As Byte
    
    Dim r As New RijndaelManaged
    r.Mode = CipherMode.CFB
    r.FeedbackSize = 8
    r.Padding = PaddingMode.None
    r.Key = CorArray.CreateInstance(vbByte, 32)
    r.IV = CorArray.CreateInstance(vbByte, 16)
    Dim Encryptor As ICryptoTransform
    Set Encryptor = r.CreateEncryptor
    EncData = Encryptor.TransformFinalBlock(Original, 0, CorArray.Length(Original))
    Assert.That CorArray.Length(EncData), Iz.EqualTo(Size), "EncData.Length"
    Assert.That EncData, Iz.EqualTo(Expected), "encrypted"
    Dim Decryptor As ICryptoTransform
    Set Decryptor = r.CreateDecryptor
    DecData = Decryptor.TransformFinalBlock(EncData, 0, CorArray.Length(EncData))
    Assert.That DecData, Iz.EqualTo(Original), "roundtrip"
End Sub

Public Sub CBC_0()
    Dim PlainText(31) As Byte
    Dim IV(15) As Byte
    Dim Key(15) As Byte
    Dim r As New RijndaelManaged
    r.BlockSize = 128
    r.Mode = CipherMode.CBC
    r.Padding = PaddingMode.Zeros
    
    Dim Expected() As Byte
    Expected = NewBytes(&H66, &HE9, &H4B, &HD4, &HEF, &H8A, &H2C, &H3B, &H88, &H4C, &HFA, &H59, &HCA, &H34, &H2B, &H2E, &HF7, &H95, &HBD, &H4A, &H52, &HE2, &H9E, &HD7, &H13, &HD3, &H13, &HFA, &H20, &HE9, &H8D, &HBC)
    
    CheckCBC r.CreateEncryptor(Key, IV), r.CreateDecryptor(Key, IV), PlainText, Expected
End Sub

Public Sub CBC_1()
    Dim PlainText(31) As Byte
    Dim IV(15) As Byte
    Dim Key(15) As Byte
    Dim i As Long
    For i = 0 To UBound(IV)
        IV(i) = i
    Next
    
    Dim r As New RijndaelManaged
    r.BlockSize = 128
    r.Mode = CipherMode.CBC
    r.Padding = PaddingMode.Zeros
    
    Dim Expected() As Byte
    Expected = NewBytes(&H7A, &HCA, &HF, &HD9, &HBC, &HD6, &HEC, &H7C, &H9F, &H97, &H46, &H66, &H16, &HE6, &HA2, &H82, &H66, &HC5, &H84, &H17, &H1D, &H3C, &H20, &H53, &H6F, &HA, &H9, &HDC, &H4D, &H1E, &H45, &H3B)
    
    CheckCBC r.CreateEncryptor(Key, IV), r.CreateDecryptor(Key, IV), PlainText, Expected
End Sub


Private Function CreateTransform(ByVal KeySize As Long, ByRef CreateMethod As String) As ICryptoTransform
    Dim IV(0 To 15) As Byte
    Dim Key() As Byte
    Key = CreateKey(KeySize)
    
    Dim r As New RijndaelManaged
    r.Mode = CipherMode.ECB
    r.Padding = PaddingMode.None
    r.KeySize = KeySize
    
    Set CreateTransform = CallByName(r, CreateMethod, VbMethod, Key, IV)
End Function

Private Function CreateKey(ByVal KeySize As Long) As Byte()
    Dim Key() As Byte
    Dim i As Long
    
    ReDim Key(0 To (KeySize \ 8) - 1)
    For i = 0 To UBound(Key)
        Key(i) = i
    Next
    
    CreateKey = Key
End Function

Private Function CreatePlainText() As Byte()
    Dim PlainText(0 To 15) As Byte
    Dim i As Long
    
    For i = 0 To 15
        PlainText(i) = i * 16 + i
    Next
    
    CreatePlainText = PlainText
End Function

Private Sub CheckCBC(ByVal Encryptor As ICryptoTransform, ByVal Decryptor As ICryptoTransform, ByRef PlainText() As Byte, ByRef Expected() As Byte)
    Dim CipherText() As Byte
    ReDim CipherText(UBound(PlainText))
    
    Dim i As Long
    For i = 0 To UBound(PlainText) Step Encryptor.InputBlockSize
        Encryptor.TransformBlock PlainText, i, Encryptor.InputBlockSize, CipherText, i
    Next
    Assert.That CipherText, Iz.EqualTo(Expected), "CBC"
    
    Dim RoundTrip() As Byte
    ReDim RoundTrip(UBound(PlainText))
    For i = 0 To UBound(CipherText) Step Decryptor.InputBlockSize
        Decryptor.TransformBlock CipherText, i, Decryptor.InputBlockSize, RoundTrip, i
    Next
    Assert.That RoundTrip, Iz.EqualTo(PlainText), "CBC-rt"
End Sub


















