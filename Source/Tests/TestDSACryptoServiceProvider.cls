VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDSACryptoServiceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDSACryptoServiceProvider
'
Option Explicit

Private FF As Long
Private PT() As Byte
Private HT() As Byte
Private Sig() As Byte
Private HA As String
Private csp As DSACryptoServiceProvider
Private G() As Byte
Private P() As Byte
Private Q() As Byte
Private X() As Byte
Private Y() As Byte
Private Seed() As Byte
Private Counter As Long



Public Sub Setup()
    Set csp = New DSACryptoServiceProvider
    Call csp.ImportCspBlob(File.ReadAllBytes(App.Path & "\Cryptography Files\DSA_Default_CspBlob.txt"))
    
    FF = FreeFile
    Open App.Path & "\Cryptography Files\DSA_Default_VerifyHash.txt" For Input As #FF
End Sub

Public Sub TearDown()
    Close #FF
End Sub



Public Sub VerifySignedHashes()
    Dim HashObj As HashAlgorithm
    
    Do While GetNextEntry
        Assert.IsTrue csp.VerifyHash(HT, "SHA1", Sig)
        
        HT(0) = (HT(0) + 1) Mod 256
        Assert.IsFalse csp.VerifyHash(HT, "SHA1", Sig)
    Loop
End Sub

Public Sub SignHashes()
    Dim HashObj As HashAlgorithm

    Do While GetNextEntry
        Dim Signature() As Byte
        Signature = csp.SignHash(HT, "sha1")
        Assert.IsTrue csp.VerifyHash(HT, "sha1", Signature)

        HT(0) = (HT(0) + 1) Mod 256
        Assert.IsFalse csp.VerifyHash(HT, "sha1", Signature)
    Loop
End Sub

Public Sub VerifyData()
    Do While GetNextEntry
        Assert.IsTrue csp.VerifyData(PT, Sig)

        PT(0) = (PT(0) + 1) Mod 256
        Assert.IsFalse csp.VerifyData(PT, Sig)
    Loop
End Sub

Public Sub VerifySignature()
    Do While GetNextEntry
        Assert.IsTrue csp.VerifySignature(HT, Sig)
        
        HT(0) = (HT(0) + 1) Mod 256
        Assert.IsFalse csp.VerifySignature(HT, Sig)
    Loop
End Sub

Public Sub SignData()
    Do While GetNextEntry
        Dim Signature() As Byte
        Signature = csp.SignData(PT)
        Assert.IsTrue csp.VerifyData(PT, Signature)

        PT(0) = (PT(0) + 1) Mod 256
        Assert.IsFalse csp.VerifyData(PT, Signature)
    Loop
End Sub

Public Sub CreateSignature()
    Do While GetNextEntry
        Dim Signature() As Byte
        Signature = csp.CreateSignature(HT)
        Assert.IsTrue csp.VerifySignature(HT, Signature)

        HT(0) = (HT(0) + 1) Mod 256
        Assert.IsFalse csp.VerifySignature(HT, Signature)
    Loop
End Sub

Public Sub RoundTripUsingCspBlob()
    Dim csp1 As New DSACryptoServiceProvider
    Dim csp2 As New DSACryptoServiceProvider
    Dim Data() As Byte
    Data = StrConv("Hello world", vbFromUnicode)
    
    Dim Sig() As Byte
    Sig = csp1.SignData(Data)
    
    Call csp2.ImportCspBlob(csp1.ExportCspBlob(True))
    Assert.IsFalse csp2.PublicOnly, "Should not be public only."
    Assert.IsTrue csp2.VerifyData(Data, Sig)
    
    Sig = csp2.SignData(Data)
    Assert.IsTrue csp1.VerifyData(Data, Sig)
    
    Data(0) = (Data(0) + 1) Mod 256
    
    Assert.IsFalse csp2.VerifyData(Data, Sig)
End Sub

Public Sub RoundTripUsingDSAParamters()
    Dim csp1 As New DSACryptoServiceProvider
    Dim csp2 As New DSACryptoServiceProvider
    Dim Data() As Byte
    Data = StrConv("Hello world", vbFromUnicode)
    
    Dim Sig() As Byte
    Sig = csp1.SignData(Data)
    
    Call csp2.ImportParameters(csp1.ExportParameters(True))
    Assert.IsFalse csp2.PublicOnly, "Should not be public only."
    
    Assert.IsTrue csp2.VerifyData(Data, Sig)
    
    Data(0) = (Data(0) + 1) Mod 256
    
    Assert.IsFalse csp2.VerifyData(Data, Sig)
End Sub

Public Sub RoundTripUsingXmlIncludePrivate()
    Dim csp1 As New DSACryptoServiceProvider
    Dim csp2 As New DSACryptoServiceProvider
    Dim Data() As Byte
    Data = StrConv("Hello world", vbFromUnicode)
    
    Dim Sig() As Byte
    Sig = csp1.SignData(Data)
    
    Call csp2.FromXmlString(csp1.ToXmlString(True))
    Assert.IsFalse csp2.PublicOnly, "Should not be public only."
    
    Assert.IsTrue csp2.VerifyData(Data, Sig)
    
    Data(0) = (Data(0) + 1) Mod 256
    
    Assert.IsFalse csp2.VerifyData(Data, Sig)
End Sub

Public Sub RoundTripUsingXmlPublicOnly()
    Dim csp1 As New DSACryptoServiceProvider
    Dim csp2 As New DSACryptoServiceProvider
    
    Call csp2.FromXmlString(csp1.ToXmlString(False))
    Assert.IsTrue csp2.PublicOnly, "Should be public only."
    
    Dim Data() As Byte
    Data = StrConv("Hello world", vbFromUnicode)
    
    Dim Sig() As Byte
    Sig = csp1.SignData(Data)
    
    Assert.IsTrue csp2.VerifyData(Data, Sig)
    
    Data(0) = (Data(0) + 1) Mod 256
    
    Assert.IsFalse csp2.VerifyData(Data, Sig)
End Sub

Public Sub CreateDsaParameters()
    FF = FreeFile
    Open App.Path & "\Cryptography Files\DSA_Default_Parameters.txt" For Input As #FF
    
    Dim line As String
    Line Input #FF, line
    Do While Len(line) > 0
        Dim Parts() As String
        Parts = Split(line, "=")
        Select Case LCase$(Trim$(Parts(0)))
            Case "g": G = ConvertHexToBytes(Trim$(Parts(1)))
            Case "p": P = ConvertHexToBytes(Trim$(Parts(1)))
            Case "q": Q = ConvertHexToBytes(Trim$(Parts(1)))
            Case "y": Y = ConvertHexToBytes(Trim$(Parts(1)))
            Case "x": X = ConvertHexToBytes(Trim$(Parts(1)))
            Case "seed": Seed = ConvertHexToBytes(Trim$(Parts(1)))
            Case "counter": Counter = Trim$(Parts(1))
        End Select
        
        Line Input #FF, line
    Loop
    Close #FF
    
    Dim csp As New DSACryptoServiceProvider
    Call csp.ImportCspBlob(File.ReadAllBytes(App.Path & "\Cryptography Files\DSA_Default_CspBlob.txt"))
    
    Dim Params As DSAParameters
    Set Params = csp.ExportParameters(True)
    
    Assert.AreEqual G, Params.G, "G"
    Assert.AreEqual P, Params.P, "P"
    Assert.AreEqual Q, Params.Q, "Q"
    Assert.AreEqual Y, Params.Y, "Y"
    Assert.AreEqual X, Params.X, "X"
    Assert.AreEqual Seed, Params.Seed, "Seed"
    Assert.AreEqual Counter, Params.Counter, "Counter"
End Sub

Public Sub TestToXmlString()
    Const PrivateXmlString As String = "<DSAKeyValue><P>gTUKJ4MTz0bpX88oKiT4M1x5/+nRGFB0rDC4QngZJMavR9U12lrH9Kibqa6CxRw0qsagHY004msAT6ZA6cWtvk9lSib2Eh/z9ZsZiSrR2RGrnedJYstaxcFUQpGk6yu/6mQb/d8MUOJt5sWq9xp+cOZTkiEkcMwAb8DFMOm9yP8=</P><Q>+i9F7Pc669JF8UiwhYRS5S1HF8E=</Q><G>BJ6y2nEAykY7qxQxQusR1PFWkVMWt4MN3709NdSURUYfKKnOVJ5F3bnvmpoUXSjjvmDp0XE1x+AuljIlpXTtSUboPBVvolTZFnxvUXW21YPzj9tGV9ZbLtOp4T2pJ3jEvaJSsw0G9ttERD37iV+4+rvEDkHrY8yxzv+9ePHC2Dg=</G><Y>e3PBHN4qPsdDY1cATb5FB4EJunc5MDmbKoarY66JRazRUJji51VqMRT8xrVjZiuySMkvbXiZyKLQypy0qaxc9NI5TUHiMDMPfzQqoxdO3Jn6F75YOlW0oFWuogBDsAM/FD1hz66om7MkJniyMfgyo6oWTv7L//lFyfiRGoyl28A=</Y><J>hDXjergw7qDf/vriWA3QZXkSms1o16ODoWM63T28s/tHZHGXHDgObylpgRvmU+NoAZR7NmVBT55VOpA0YeV2SWw8p8oqrCVrqTb1X1BiQ5RoRW0klVHtEig4srpyye9RYWsEMCnnmBVoBhh+</J><Seed>6/momY8s5RGixos8k9K3j63pag0=</Seed><PgenCounter>nQ==</PgenCounter><X>IadKN7otvgvWVGhGAVACIsQ+E0c=</X></DSAKeyValue>"
    Const PublicXmlSTring As String = "<DSAKeyValue><P>gTUKJ4MTz0bpX88oKiT4M1x5/+nRGFB0rDC4QngZJMavR9U12lrH9Kibqa6CxRw0qsagHY004msAT6ZA6cWtvk9lSib2Eh/z9ZsZiSrR2RGrnedJYstaxcFUQpGk6yu/6mQb/d8MUOJt5sWq9xp+cOZTkiEkcMwAb8DFMOm9yP8=</P><Q>+i9F7Pc669JF8UiwhYRS5S1HF8E=</Q><G>BJ6y2nEAykY7qxQxQusR1PFWkVMWt4MN3709NdSURUYfKKnOVJ5F3bnvmpoUXSjjvmDp0XE1x+AuljIlpXTtSUboPBVvolTZFnxvUXW21YPzj9tGV9ZbLtOp4T2pJ3jEvaJSsw0G9ttERD37iV+4+rvEDkHrY8yxzv+9ePHC2Dg=</G><Y>e3PBHN4qPsdDY1cATb5FB4EJunc5MDmbKoarY66JRazRUJji51VqMRT8xrVjZiuySMkvbXiZyKLQypy0qaxc9NI5TUHiMDMPfzQqoxdO3Jn6F75YOlW0oFWuogBDsAM/FD1hz66om7MkJniyMfgyo6oWTv7L//lFyfiRGoyl28A=</Y><J>hDXjergw7qDf/vriWA3QZXkSms1o16ODoWM63T28s/tHZHGXHDgObylpgRvmU+NoAZR7NmVBT55VOpA0YeV2SWw8p8oqrCVrqTb1X1BiQ5RoRW0klVHtEig4srpyye9RYWsEMCnnmBVoBhh+</J><Seed>6/momY8s5RGixos8k9K3j63pag0=</Seed><PgenCounter>nQ==</PgenCounter></DSAKeyValue>"
    
    Dim csp As New DSACryptoServiceProvider
    Call csp.ImportCspBlob(File.ReadAllBytes(App.Path & "\Cryptography Files\DSA_Default_CspBlob.txt"))
    Debug.Print csp.ToXmlString(False)
    
    Assert.AreEqual PublicXmlSTring, csp.ToXmlString(False)
    Assert.AreEqual PrivateXmlString, csp.ToXmlString(True)
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
            Case "ht": HT = ConvertHexToBytes(Trim$(Parts(1)))
            Case "ha": HA = Trim$(Parts(1))
            Case "sig": Sig = ConvertHexToBytes(Trim$(Parts(1)))
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


