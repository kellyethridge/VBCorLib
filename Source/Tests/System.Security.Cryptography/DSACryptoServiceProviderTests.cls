VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSACryptoServiceProviderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' DSACryptoServiceProvider
'
Option Explicit
Implements ICategorizable

Private Const PrivateCspBlob As String = "0702 0000 0022 0000 4453 5332 0004 0000 ffc8 bde9 30c5 c06f 00cc 7024 2192 53e6 707e 1af7 aac5 e66d e250 0cdf fd1b 64ea bf2b eba4 9142 54c1 c55a cb62 49e7 9dab 11d9 d12a 8919 9bf5 f31f 12f6 264a 654f bead c5e9 40a6 4f00 6be2 348d 1da0 c6aa 341c c582 aea9 9ba8 f4c7 5ada 35d5 47af c624 1978 42b8 30ac 7450 18d1 e9ff 795c 33f8 242a 28cf 5fe9 46cf 1383 270a 3581 c117 472d e552 8485 b048 f145 d2eb 3af7 ec45 2ffa 38d8 c2f1 78bd ffce b1cc 63eb 410e c4bb fab8 5f89 fb3d 4444 dbf6 060d b352 a2bd c478 27a9 3de1 a9d3 2e5b d657 46db 8ff3 83d5 b675 516f 7c16 d954 a26f 153c e846 49ed 74a5 2532 962e e0c7 3571 d1e9 60be e328 5d14 9a9a efb9 dd45 9e54 cea9 281f 4645 94d4 353d bddf 0d83 b716 5391 56f1 d411 eb42 3114 ab3b 46ca 0071 dab2 9e04 4713 3ec4 2202 5001 4668 54d6 0bbe 2dba 374a a721 9d00 0000 0d6a e9ad 8fb7 d293 3c8b c6a2 11e5 2c8f 99a8 f9eb"
Private Const PublicCspBlob As String = "0602 0000 0022 0000 4453 5331 0004 0000 ffc8 bde9 30c5 c06f 00cc 7024 2192 53e6 707e 1af7 aac5 e66d e250 0cdf fd1b 64ea" & "bf2b eba4 9142 54c1 c55a cb62 49e7 9dab" & "11d9 d12a 8919 9bf5 f31f 12f6 264a 654f bead c5e9 40a6 4f00 6be2 348d 1da0 c6aa" & "341c c582 aea9 9ba8 f4c7 5ada 35d5 47af" & "c624 1978 42b8 30ac 7450 18d1 e9ff 795c 33f8 242a 28cf 5fe9 46cf 1383 270a 3581" & "c117 472d e552 8485 b048 f145 d2eb 3af7" & "ec45 2ffa 38d8 c2f1 78bd ffce b1cc 63eb 410e c4bb fab8 5f89 fb3d 4444 dbf6 060d b352 a2bd c478 27a9 3de1 a9d3 2e5b d657" & _
                                        "46db 8ff3 83d5 b675 516f 7c16 d954 a26f 153c e846 49ed 74a5 2532 962e e0c7 3571 d1e9 60be e328 5d14 9a9a efb9 dd45 9e54 cea9 281f 4645 94d4 353d bddf 0d83 b716 5391 56f1 d411 eb42 3114 ab3b 46ca 0071 dab2 9e04 c0db a58c 1a91 f8c9 45f9 ffcb fe4e 16aa a332 f831 b278 2624 b39b a8ae cf61 3d14 3f03 b043 00a2 ae55 a0b4 553a 58be 17fa 99dc 4e17 a32a 347f 0f33 30e2 414d 39d2 f45c aca9 b49c cad0 a2c8 9978 6d2f c948 b22b 6663 b5c6 fc14 316a 55e7 e298 50d1 ac45 89ae 63ab 862a 9b39 3039 77ba 0981 0745 be4d 0057 6343 c73e 2ade 1cc1 737b 9d00 0000 0d6a e9ad 8fb7 d293 3c8b c6a2 11e5 2c8f 99a8 f9eb"
Private Const PublicXmlString As String = "<DSAKeyValue><P>gTUKJ4MTz0bpX88oKiT4M1x5/+nRGFB0rDC4QngZJMavR9U12lrH9Kibqa6CxRw0qsagHY004msAT6ZA6cWtvk9lSib2Eh/z9ZsZiSrR2RGrnedJYstaxcFUQpGk6yu/6mQb/d8MUOJt5sWq9xp+cOZTkiEkcMwAb8DFMOm9yP8=</P><Q>+i9F7Pc669JF8UiwhYRS5S1HF8E=</Q><G>BJ6y2nEAykY7qxQxQusR1PFWkVMWt4MN3709NdSURUYfKKnOVJ5F3bnvmpoUXSjjvmDp0XE1x+AuljIlpXTtSUboPBVvolTZFnxvUXW21YPzj9tGV9ZbLtOp4T2pJ3jEvaJSsw0G9ttERD37iV+4+rvEDkHrY8yxzv+9ePHC2Dg=</G><Y>e3PBHN4qPsdDY1cATb5FB4EJunc5MDmbKoarY66JRazRUJji51VqMRT8xrVjZiuySMkvbXiZyKLQypy0qaxc9NI5TUHiMDMPfzQqoxdO3Jn6F75YOlW0oFWuogBDsAM/FD1hz66om7MkJniyMfgyo6oWTv7L//lFyfiRGoyl28A=</Y><J>hDXjergw7qDf/vriWA3QZXkSms1o16ODoWM63T28s/tHZHGXHDgObylpgRvmU+NoAZR7NmVBT55VOpA0YeV2SWw8p8oqrCVrqTb1X1BiQ5RoRW0klVHtEig4srpyye9RYWsEMCnnmBVoBhh+</J><Seed>6/momY8s5RGixos8k9K3j63pag0=</Seed><PgenCounter>nQ==</PgenCounter></DSAKeyValue>"
Private Const PrivateXmlString As String = "<DSAKeyValue><P>gTUKJ4MTz0bpX88oKiT4M1x5/+nRGFB0rDC4QngZJMavR9U12lrH9Kibqa6CxRw0qsagHY004msAT6ZA6cWtvk9lSib2Eh/z9ZsZiSrR2RGrnedJYstaxcFUQpGk6yu/6mQb/d8MUOJt5sWq9xp+cOZTkiEkcMwAb8DFMOm9yP8=</P><Q>+i9F7Pc669JF8UiwhYRS5S1HF8E=</Q><G>BJ6y2nEAykY7qxQxQusR1PFWkVMWt4MN3709NdSURUYfKKnOVJ5F3bnvmpoUXSjjvmDp0XE1x+AuljIlpXTtSUboPBVvolTZFnxvUXW21YPzj9tGV9ZbLtOp4T2pJ3jEvaJSsw0G9ttERD37iV+4+rvEDkHrY8yxzv+9ePHC2Dg=</G><Y>e3PBHN4qPsdDY1cATb5FB4EJunc5MDmbKoarY66JRazRUJji51VqMRT8xrVjZiuySMkvbXiZyKLQypy0qaxc9NI5TUHiMDMPfzQqoxdO3Jn6F75YOlW0oFWuogBDsAM/FD1hz66om7MkJniyMfgyo6oWTv7L//lFyfiRGoyl28A=</Y><J>hDXjergw7qDf/vriWA3QZXkSms1o16ODoWM63T28s/tHZHGXHDgObylpgRvmU+NoAZR7NmVBT55VOpA0YeV2SWw8p8oqrCVrqTb1X1BiQ5RoRW0klVHtEig4srpyye9RYWsEMCnnmBVoBhh+</J><Seed>6/momY8s5RGixos8k9K3j63pag0=</Seed><PgenCounter>nQ==</PgenCounter><X>IadKN7otvgvWVGhGAVACIsQ+E0c=</X></DSAKeyValue>"

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub KeyExchangeAlgorithm_WhenCalled_ReturnsEmpty()
    Dim Csp As New DSACryptoServiceProvider
    
    Dim Actual As String
    Actual = Csp.KeyExchangeAlgorithm
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub LegalKeySizes_WhenCalled_ReturnsListOfKeySizes()
    Dim Csp As New DSACryptoServiceProvider
    Dim Actual() As KeySizes
    
    Actual = Csp.LegalKeySizes
    
    Assert.That CorArray.Length(Actual), Iz.EqualTo(1)
    Assert.That Actual(0).MinSize, Iz.EqualTo(512)
    Assert.That Actual(0).MaxSize, Iz.EqualTo(1024)
    Assert.That Actual(0).SkipSize, Iz.EqualTo(64)
End Sub

Public Sub ExportParameters_WithFalse_ExcludesPrivateXParameter()
    Dim Parameters  As DSAParameters
    Dim Csp         As New DSACryptoServiceProvider
    Dim Actual      As DSAParameters
    
    Set Parameters = MakePrivateParameters
    
    Csp.ImportParameters Parameters
    Set Actual = Csp.ExportParameters(False)
    
    Assert.That Actual.p, Iz.EqualTo(Parameters.p), "#1"
    Assert.That Actual.Q, Iz.EqualTo(Parameters.Q), "#2"
    Assert.That Actual.g, Iz.EqualTo(Parameters.g), "#3"
    Assert.That Actual.j, Iz.EqualTo(Parameters.j), "#4"
    Assert.That Actual.y, Iz.EqualTo(Parameters.y), "#5"
    Assert.That Actual.x, Iz.Empty, "#6"
    Assert.That Actual.Seed, Iz.EqualTo(Parameters.Seed), "#7"
    Assert.That Actual.Counter, Iz.EqualTo(Parameters.Counter), "#8"
End Sub

Public Sub ExportParameters_WithTrue_IncludesPrivateXParameter()
    Dim Parameters  As DSAParameters
    Dim Csp         As New DSACryptoServiceProvider
    Dim Actual      As DSAParameters
    
    Set Parameters = MakePrivateParameters
    
    Csp.ImportParameters Parameters
    Set Actual = Csp.ExportParameters(True)
    
    Assert.That Actual.p, Iz.EqualTo(Parameters.p), "#1"
    Assert.That Actual.Q, Iz.EqualTo(Parameters.Q), "#2"
    Assert.That Actual.g, Iz.EqualTo(Parameters.g), "#3"
    Assert.That Actual.j, Iz.EqualTo(Parameters.j), "#4"
    Assert.That Actual.y, Iz.EqualTo(Parameters.y), "#5"
    Assert.That Actual.x, Iz.EqualTo(Parameters.x), "#6"
    Assert.That Actual.Seed, Iz.EqualTo(Parameters.Seed), "#7"
    Assert.That Actual.Counter, Iz.EqualTo(Parameters.Counter), "#8"
End Sub

Public Sub ExportParameters_WithTrueAndIsOnlyPublic_ThrowsCryptographicException()
    Dim Csp As New DSACryptoServiceProvider
    Csp.ImportParameters MakePublicParameters
    
    On Error Resume Next
    Csp.ExportParameters True
    
    AssertCryptographicException Err
End Sub

Public Sub ToXmlString_WithPublicParametersOnly_ReturnsPublicXmlString()
    Dim Csp     As New DSACryptoServiceProvider
    Dim Actual  As SecurityElement
    
    Csp.FromXmlString PublicXmlString
    Set Actual = SecurityElement.FromString(Csp.ToXmlString(False))
    
    Dim Expected As SecurityElement
    Set Expected = SecurityElement.FromString(PublicXmlString)
    Assert.That Actual.Equal(Expected), Iz.True
End Sub

Public Sub ToXmlString_WithPublicParametersOnlyAndUsesPrivateParameters_ReturnsPublicXmlString()
    Dim Csp     As New DSACryptoServiceProvider
    Dim Actual  As SecurityElement
    
    Csp.FromXmlString PrivateXmlString
    Set Actual = SecurityElement.FromString(Csp.ToXmlString(False))
    
    Dim Expected As SecurityElement
    Set Expected = SecurityElement.FromString(PublicXmlString)
    Assert.That Actual.Equal(Expected), Iz.True
End Sub

Public Sub ToXmlString_WithPrivateParameters_ReturnsPrivateXmlString()
    Dim Csp     As New DSACryptoServiceProvider
    Dim Actual  As SecurityElement
    
    Csp.FromXmlString PrivateXmlString
    Set Actual = SecurityElement.FromString(Csp.ToXmlString(True))
    
    Dim Expected As SecurityElement
    Set Expected = SecurityElement.FromString(PrivateXmlString)
    Assert.That Actual.Equal(Expected), Iz.True
End Sub

Public Sub CspBlob_WithImportExport_RoundTrips()
    Dim Blob()      As Byte
    Dim Csp         As New DSACryptoServiceProvider
    Dim Actual()    As Byte
    
    Blob = MakePrivateCspBlob
    Csp.ImportCspBlob Blob
    Actual = Csp.ExportCspBlob(True)
    
    Assert.That Actual, Iz.EqualTo(Blob)
End Sub

Public Sub ImportCspBlob_WithNullArray_ThrowsArgumentNullException()
    Dim Csp As New DSACryptoServiceProvider
    
    On Error Resume Next
    Csp.ImportCspBlob NullBytes
    
    AssertArgumentNullException Err, "KeyBlob"
End Sub

Public Sub PublicOnly_WithPrivateCspBlob_ReturnsFalse()
    Dim Csp     As New DSACryptoServiceProvider
    Dim Actual  As Boolean
    Csp.ImportCspBlob MakePrivateCspBlob
    
    Actual = Csp.PublicOnly
    
    Assert.That Actual, Iz.False
End Sub

Public Sub PublicOnly_WithPublicCspBlob_ReturnsTrue()
    Dim Csp     As New DSACryptoServiceProvider
    Dim Actual  As Boolean
    Csp.ImportCspBlob MakePublicCspBlob
    
    Actual = Csp.PublicOnly
    
    Assert.That Actual, Iz.True
End Sub

Public Sub PublicOnly_WithPrivateParameters_ReturnsFalse()
    Dim Csp     As New DSACryptoServiceProvider
    Dim Actual  As Boolean
    Csp.ImportParameters MakePrivateParameters
    
    Actual = Csp.PublicOnly
    
    Assert.That Actual, Iz.False
End Sub

Public Sub PublicOnly_WithPublicParameters_ReturnsTrue()
    Dim Csp     As New DSACryptoServiceProvider
    Dim Actual  As Boolean
    Csp.ImportParameters MakePublicParameters
    
    Actual = Csp.PublicOnly
    
    Assert.That Actual, Iz.True
End Sub

Public Sub ImportCspBlob_WithPrivateTypeAndPublicMagicNumber_ThrowsCryptographicException()
    Dim Csp As New DSACryptoServiceProvider
    Dim BadData() As Byte
    BadData = MakePublicCspBlob
    BadData(0) = PRIVATEKEYBLOB
    
    On Error Resume Next
    Csp.ImportCspBlob BadData
    
    AssertCryptographicException Err
End Sub

Public Sub ImportCspBlob_WithPublicTypeAndPrivateMagicNumber_ThrowsCryptographicException()
    Dim Csp As New DSACryptoServiceProvider
    Dim BadData() As Byte
    BadData = MakePrivateCspBlob
    BadData(0) = PUBLICKEYBLOB
    
    On Error Resume Next
    Csp.ImportCspBlob BadData
    
    AssertCryptographicException Err
End Sub

Public Sub VerifySignature_WithValidSignature_ReturnsTrue()
    Dim Csp             As New DSACryptoServiceProvider
    Dim Hash(0 To 19)   As Byte
    Dim Signature()     As Byte
    Dim Actual          As Boolean
    Csp.ImportCspBlob MakePrivateCspBlob
    
    Signature = Csp.CreateSignature(Hash)
    Actual = Csp.VerifySignature(Hash, Signature)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub VerifySignature_WithInvalidSignature_ReturnsFalse()
    Dim Csp             As DSACryptoServiceProvider
    Dim Hash(0 To 19)   As Byte
    Dim Signature()     As Byte
    Dim Actual          As Boolean
    Set Csp = MakeSignableCsp
    
    Signature = Csp.CreateSignature(Hash)
    Signature(0) = Signature(0) Xor &HFF
    Actual = Csp.VerifySignature(Hash, Signature)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub CreateSignature_WithNullHash_ThrowsArgumentNullException()
    Dim Csp As New DSACryptoServiceProvider
    
    On Error Resume Next
    Csp.CreateSignature NullBytes
    
    AssertArgumentNullException Err, "RgbHash"
End Sub

Public Sub VerifySignature_WithNullHash_ThrowsArgumentNullException()
    Dim Csp As New DSACryptoServiceProvider
    Dim Signature(0 To 39) As Byte
    
    On Error Resume Next
    Csp.VerifySignature NullBytes, Signature
    
    AssertArgumentNullException Err, "RgbHash"
End Sub

Public Sub VerifyHash_WithSignedHash_ReturnsTrue()
    Dim Csp As DSACryptoServiceProvider
    Dim Hash(0 To 19) As Byte
    Dim Signature() As Byte
    Dim Actual As Boolean
    Set Csp = MakeSignableCsp
    
    Signature = Csp.SignHash(Hash, "sha")
    Actual = Csp.VerifyHash(Hash, "sha", Signature)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub VerifyHash_WithInvalidSignature_ReturnsFalse()
    Dim Csp             As DSACryptoServiceProvider
    Dim Hash(0 To 19)   As Byte
    Dim Signature()     As Byte
    Dim Actual          As Boolean
    Set Csp = MakeSignableCsp
        
    Signature = Csp.SignHash(Hash, "sha")
    Signature(0) = Signature(0) Xor &HFF
    Actual = Csp.VerifyHash(Hash, "sha", Signature)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub VerifyHash_WithSignatureNot40Bytes_ThrowsCryptographicException()
    Dim Csp As New DSACryptoServiceProvider
    Dim Hash(0 To 19) As Byte
    Dim Signature(0) As Byte
    
    On Error Resume Next
    Csp.VerifyHash Hash, "sha", Signature
    
    AssertCryptographicException Err
End Sub

Private Function MakeSignableCsp() As DSACryptoServiceProvider
    Set MakeSignableCsp = New DSACryptoServiceProvider
    MakeSignableCsp.ImportCspBlob MakePrivateCspBlob
End Function

Private Function MakePrivateParameters() As DSAParameters
    Dim Parameters As DSAParameters
    Set Parameters = MakePublicParameters
    Parameters.x = HexToBytes("21A74A37BA2DBE0BD654684601500222C43E1347")
    
    Set MakePrivateParameters = Parameters
End Function

Private Function MakePublicParameters() As DSAParameters
    Dim Parameters As New DSAParameters
    Parameters.p = HexToBytes("81350A278313CF46E95FCF282A24F8335C79FFE9D1185074AC30B842781924C6AF47D535DA5AC7F4A89BA9AE82C51C34AAC6A01D8D34E26B004FA640E9C5ADBE4F654A26F6121FF3F59B19892AD1D911AB9DE74962CB5AC5C1544291A4EB2BBFEA641BFDDF0C50E26DE6C5AAF71A7E70E65392212470CC006FC0C530E9BDC8FF")
    Parameters.Q = HexToBytes("FA2F45ECF73AEBD245F148B0858452E52D4717C1")
    Parameters.g = HexToBytes("049EB2DA7100CA463BAB143142EB11D4F156915316B7830DDFBD3D35D49445461F28A9CE549E45DDB9EF9A9A145D28E3BE60E9D17135C7E02E963225A574ED4946E83C156FA254D9167C6F5175B6D583F38FDB4657D65B2ED3A9E13DA92778C4BDA252B30D06F6DB44443DFB895FB8FABBC40E41EB63CCB1CEFFBD78F1C2D838")
    Parameters.y = HexToBytes("7B73C11CDE2A3EC7436357004DBE45078109BA773930399B2A86AB63AE8945ACD15098E2E7556A3114FCC6B563662BB248C92F6D7899C8A2D0CA9CB4A9AC5CF4D2394D41E230330F7F342AA3174EDC99FA17BE583A55B4A055AEA20043B0033F143D61CFAEA89BB3242678B231F832A3AA164EFECBFFF945C9F8911A8CA5DBC0")
    Parameters.Seed = HexToBytes("EBF9A8998F2CE511A2C68B3C93D2B78FADE96A0D")
    Parameters.j = HexToBytes("21A74A37BA2DBE0BD654684601500222C43E1347")
    Parameters.Counter = 157
    
    Set MakePublicParameters = Parameters
End Function

Private Function MakePrivateCspBlob() As Byte()
    MakePrivateCspBlob = HexToBytes(PrivateCspBlob)
End Function

Private Function MakePublicCspBlob() As Byte()
    MakePublicCspBlob = HexToBytes(PublicCspBlob)
End Function

