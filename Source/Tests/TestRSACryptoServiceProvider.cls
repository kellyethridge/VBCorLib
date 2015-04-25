VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRSACryptoServiceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRSACryptoServiceProvider
'
Option Explicit
Implements ICategorizable

Private FF As Long
Private PT() As Byte
Private CT() As Byte
Private d() As Byte
Private Q() As Byte
Private P() As Byte
Private DP() As Byte
Private DQ() As Byte
Private Modulus() As Byte
Private Exponent() As Byte
Private InverseQ() As Byte
Private Sig() As Byte
Private HA As String

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub



Public Sub Setup()
    FF = FreeFile
    Open App.Path & "\Cryptography Files\RSA_Default_Entries.txt" For Input As #FF
End Sub

Public Sub Teardown()
    Close #FF
End Sub

Public Sub CreateDefault()
    Dim csp As New RSACryptoServiceProvider
    Assert.That csp.KeySize, Iz.EqualTo(512).OrElse.EqualTo(1024)
End Sub

Public Sub RoundTripEncrypt()
    Dim csp As New RSACryptoServiceProvider
    Dim b() As Byte
    Dim d() As Byte
    
    d = StrConv("Hello World!", vbFromUnicode)
    b = csp.Encrypt(d, False)
    
    Erase d
    d = csp.Decrypt(b, False)
    
    Assert.AreEqual "Hello World!", StrConv(d, vbUnicode)
End Sub

Public Sub TestDecryptDefaultEntries()
    Dim csp As New RSACryptoServiceProvider
    csp.ImportCspBlob File.ReadAllBytes(App.Path & "\Cryptography Files\RSA_Default_CspBlob.txt")
    
    Dim b() As Byte
    Do While GetNextEntry
        b = csp.Decrypt(CT, False)
        Assert.AreEqual PT, b
    Loop
End Sub

Public Sub TestPublicOnly()
    Dim csp As New RSACryptoServiceProvider
    Assert.IsFalse csp.PublicOnly
    
    csp.ImportCspBlob File.ReadAllBytes(App.Path & "\Cryptography Files\RSA_Default_PublicOnly_CspBlob.txt")
    Assert.IsTrue csp.PublicOnly
    
    csp.ImportCspBlob File.ReadAllBytes(App.Path & "\Cryptography Files\RSA_Default_CspBlob.txt")
    Assert.IsFalse csp.PublicOnly
End Sub

Public Sub RoundTripUsingCspBlob()
    Dim csp1 As New RSACryptoServiceProvider
    Dim csp2 As New RSACryptoServiceProvider
    
    Dim b() As Byte
    b = csp1.Encrypt(Encoding.ASCII.GetBytes("Hello There"), True)
    
    Call csp2.ImportCspBlob(csp1.ExportCspBlob(True))
    Assert.IsFalse csp2.PublicOnly, "Should not be public only."
    
    Dim b2() As Byte
    b2 = csp2.Decrypt(b, True)
    
    Assert.AreEqual "Hello There", Encoding.ASCII.GetString(b2)
End Sub

Public Sub RoundTripPublicOnly()
    Dim csp1 As New RSACryptoServiceProvider
    Dim csp2 As New RSACryptoServiceProvider
    
    Call csp2.ImportCspBlob(csp1.ExportCspBlob(False))
    Assert.IsTrue csp2.PublicOnly, "Should be public only."
    
    Dim b() As Byte
    b = csp2.Encrypt(Encoding.ASCII.GetBytes("One Way"), False)
    
    Dim b2() As Byte
    b2 = csp1.Decrypt(b, False)
    
    Assert.AreEqual "One Way", Encoding.ASCII.GetString(b2)
End Sub

Public Sub RoundTripUsingXmlIncludePrivate()
    Dim csp1 As New RSACryptoServiceProvider
    Dim csp2 As New RSACryptoServiceProvider
    
    Dim b() As Byte
    b = csp1.Encrypt(Encoding.ASCII.GetBytes("Hello There"), False)
    
    Call csp2.FromXmlString(csp1.ToXmlString(True))
    Assert.IsFalse csp2.PublicOnly, "Should not be public only."
    
    Dim b2() As Byte
    b2 = csp2.Decrypt(b, False)
    
    Assert.AreEqual "Hello There", Encoding.ASCII.GetString(b2)
End Sub

Public Sub RoundTripXmlPublicOnly()
    Dim csp1 As New RSACryptoServiceProvider
    Dim csp2 As New RSACryptoServiceProvider
    
    
    Call csp2.FromXmlString(csp1.ToXmlString(False))
    Assert.IsTrue csp2.PublicOnly, "Should be public only."
    
    Dim b() As Byte
    b = csp2.Encrypt(Encoding.ASCII.GetBytes("Hello There"), False)
    
    Dim b2() As Byte
    b2 = csp1.Decrypt(b, False)
    
    Assert.AreEqual "Hello There", Encoding.ASCII.GetString(b2)
End Sub

Public Sub CreateRsaParameters()
    FF = FreeFile
    Open App.Path & "\Cryptography Files\RSA_Default_Parameters.txt" For Input As #FF
    
    Dim line As String
    Line Input #FF, line
    Do While Len(line) > 0
        Dim Parts() As String
        Parts = Split(line, "=")
        Select Case LCase$(Trim$(Parts(0)))
            Case "d": d = ConvertHexToBytes(Trim$(Parts(1)))
            Case "p": P = ConvertHexToBytes(Trim$(Parts(1)))
            Case "q": Q = ConvertHexToBytes(Trim$(Parts(1)))
            Case "dp": DP = ConvertHexToBytes(Trim$(Parts(1)))
            Case "dq": DQ = ConvertHexToBytes(Trim$(Parts(1)))
            Case "modulus": Modulus = ConvertHexToBytes(Trim$(Parts(1)))
            Case "exponent": Exponent = ConvertHexToBytes(Trim$(Parts(1)))
            Case "inverseq": InverseQ = ConvertHexToBytes(Trim$(Parts(1)))
        End Select
        
        Line Input #FF, line
    Loop
    Close #FF
    
    Dim csp As New RSACryptoServiceProvider
    Call csp.ImportCspBlob(File.ReadAllBytes(App.Path & "\Cryptography Files\RSA_Default_CspBlob.txt"))
    
    Dim Params As RSAParameters
    Set Params = csp.ExportParameters(True)
    
    Assert.AreEqual d, Params.d, "D"
    Assert.AreEqual P, Params.P, "P"
    Assert.AreEqual Q, Params.Q, "Q"
    Assert.AreEqual DP, Params.DP, "DP"
    Assert.AreEqual DQ, Params.DQ, "DQ"
    Assert.AreEqual Modulus, Params.Modulus, "Modulus"
    Assert.AreEqual Exponent, Params.Exponent, "Exponent"
    Assert.AreEqual InverseQ, Params.InverseQ, "InverseQ"
End Sub

Public Sub ImportRsaParameters()
    Dim csp1 As New RSACryptoServiceProvider
    Call csp1.ImportCspBlob(File.ReadAllBytes(App.Path & "\Cryptography Files\RSA_Default_CspBlob.txt"))
    
    Dim csp2 As New RSACryptoServiceProvider
    Call csp2.ImportParameters(csp1.ExportParameters(True))
    
    Dim b() As Byte
    b = csp1.Encrypt(Encoding.ASCII.GetBytes("Wow, cool"), False)
    
    Dim b2() As Byte
    b2 = csp2.Decrypt(b, False)
    
    Assert.AreEqual "Wow, cool", Encoding.ASCII.GetString(b2)
End Sub

Public Sub ExportTheImportRsaParameters()
    Dim Data() As Byte
    Data = Cor.NewBytes(5, 4, 3, 2, 1)
    
    Dim Bytes() As Byte
    
    Dim csp1 As New RSACryptoServiceProvider
    Bytes = csp1.Encrypt(Data, True)
    
    Dim Params As RSAParameters
    Set Params = csp1.ExportParameters(True)
    
    Dim csp2 As New RSACryptoServiceProvider
    Call csp2.ImportParameters(Params)
    
    Bytes = csp2.Decrypt(Bytes, True)
    
    Assert.AreEqual Data, Bytes
End Sub

Public Sub CreateUsingInvalidCspParametersShouldFail()
    Dim P As CspParameters
    Set P = Cor.NewCspParameters(13, "Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider")
    P.KeyNumber = 2
    
    Dim csp As RSACryptoServiceProvider
    Set csp = Cor.NewRSACryptoServiceProvider(, P)
    
    On Error GoTo errTrap
    Debug.Print csp.CspKeyContainerInfo.ProviderName
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is CryptographicException, "Wrong exception type caught."
End Sub

Public Sub InvalidKeyGeneration()
    ' Key generation does not occur until it is needed.
    Dim csp As RSACryptoServiceProvider
    Set csp = Cor.NewRSACryptoServiceProvider(16332)
    
    On Error GoTo errTrap
    ' This will cause the key to be generated.
    Call csp.Encrypt(Cor.NewBytes(0), False)
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is CryptographicException, "Wrong exception type caught."
End Sub

Public Sub UseAnExistingKey()
    Dim csp1 As New RSACryptoServiceProvider
    Dim PT() As Byte
    PT = Cor.NewBytes(5, 4, 3, 2, 1)
    Dim CT() As Byte
    CT = csp1.Encrypt(PT, False)
    
    Dim CspParams As CspParameters
    Set CspParams = Cor.NewCspParameters(1, csp1.CspKeyContainerInfo.ProviderName, csp1.CspKeyContainerInfo.KeyContainerName)
    CspParams.Flags = UseExistingKey
    CspParams.KeyNumber = 1
    
    csp1.PersistKeyInCsp = True
    Set csp1 = Nothing
    
    Dim csp2 As RSACryptoServiceProvider
    Set csp2 = Cor.NewRSACryptoServiceProvider(, CspParams)
    csp2.PersistKeyInCsp = False
    
    Assert.AreEqual PT, csp2.Decrypt(CT, False)
End Sub





Private Function GetNextEntry() As Boolean
    If EOF(FF) Then Exit Function
    
    Dim line As String
    Line Input #FF, line
    Do While Len(line) > 0
        Dim Parts() As String
        Parts = Split(line, "=")
        Select Case LCase$(Trim$(Parts(0)))
            Case "pt": PT = ConvertHexToBytes(Trim$(Parts(1)))
            Case "ct": CT = ConvertHexToBytes(Trim$(Parts(1)))
            Case "d": d = ConvertHexToBytes(Trim$(Parts(1)))
            Case "p": P = ConvertHexToBytes(Trim$(Parts(1)))
            Case "q": Q = ConvertHexToBytes(Trim$(Parts(1)))
            Case "dp": DP = ConvertHexToBytes(Trim$(Parts(1)))
            Case "dq": DQ = ConvertHexToBytes(Trim$(Parts(1)))
            Case "modulus": Modulus = ConvertHexToBytes(Trim$(Parts(1)))
            Case "exponent": Exponent = ConvertHexToBytes(Trim$(Parts(1)))
            Case "inverseq": InverseQ = ConvertHexToBytes(Trim$(Parts(1)))
            Case "ha": HA = Trim$(Parts(1))
            Case "sig": Sig = ConvertHexToBytes(Trim$(Parts(1)))
            Case "pt": PT = ConvertHexToBytes(Trim$(Parts(1)))
        End Select
        
        Line Input #FF, line
    Loop
    GetNextEntry = True
End Function

Private Function ConvertHexToBytes(ByVal s As String) As Byte()
    Dim ret() As Byte
    ReDim ret(0 To Len(s) \ 2 - 1)
    
    Dim i As Long
    For i = 0 To UBound(ret)
        ret(i) = CByte("&h" & Mid$(s, (i * 2) + 1, 2))
    Next i
    
    ConvertHexToBytes = ret
End Function

Private Function ConvertBytesToHex(ByRef Bytes() As Byte) As String
    Dim i As Long
    Dim sb As New StringBuilder
    
    For i = 0 To UBound(Bytes)
        sb.AppendFormat "{0:X2}", Bytes(i)
    Next i
    
    ConvertBytesToHex = sb.ToString
End Function
