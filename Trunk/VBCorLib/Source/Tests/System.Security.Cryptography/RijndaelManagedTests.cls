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





















