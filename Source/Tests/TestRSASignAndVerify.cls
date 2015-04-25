VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRSASignAndVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRSASignAndVerify
'
Option Explicit
Implements ICategorizable

Private FF As Long
Private PT() As Byte
Private ht() As Byte
Private Sig() As Byte
Private HA As String
Private csp As RSACryptoServiceProvider

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub



Public Sub Setup()
    Set csp = New RSACryptoServiceProvider
    Call csp.ImportCspBlob(File.ReadAllBytes(App.Path & "\Cryptography Files\RSA_Default_Sign_CspBlob.txt"))
    
    FF = FreeFile
    Open App.Path & "\Cryptography Files\RSA_Default_VerifyHash.txt" For Input As #FF
End Sub

Public Sub Teardown()
    Close #FF
End Sub



Public Sub VerifySignedHashes()
    Dim HashObj As HashAlgorithm
    
    Do While GetNextEntry
        Set HashObj = CryptoConfig.CreateFromName(HA)
        
        Dim hash() As Byte
        hash = HashObj.ComputeHash(PT)
        Assert.AreEqual ht, hash
        
        Assert.IsTrue csp.VerifyHash(hash, CryptoConfig.MapNameToOID(HA), Sig)
        
        hash(0) = (hash(0) + 1) Mod 256
        Assert.IsFalse csp.VerifyHash(hash, CryptoConfig.MapNameToOID(HA), Sig)
    Loop
End Sub

Public Sub SignHashes()
    Dim HashObj As HashAlgorithm
    
    Do While GetNextEntry
        Set HashObj = CryptoConfig.CreateFromName(HA)
        
        Dim hash() As Byte
        hash = HashObj.ComputeHash(PT)
        Assert.AreEqual ht, hash
        
        Dim Signature() As Byte
        Signature = csp.SignHash(hash, CryptoConfig.MapNameToOID(HA))
        Assert.AreEqual Sig, Signature
        
        hash(0) = (hash(0) + 1) Mod 256
        Signature = csp.SignHash(hash, CryptoConfig.MapNameToOID(HA))
        Assert.That Sig, Iz.Not.EqualTo(Signature)
    Loop
End Sub

Public Sub VerifyData()
    Do While GetNextEntry
        Assert.IsTrue csp.VerifyData(PT, HA, Sig)
        
        PT(0) = (PT(0) + 1) Mod 256
        Assert.IsFalse csp.VerifyData(PT, HA, Sig)
    Loop
End Sub

Public Sub SignData()
    Do While GetNextEntry
        Dim Signature() As Byte
        Signature = csp.SignData(PT, HA)
        Assert.AreEqual Sig, Signature
    
        PT(0) = (PT(0) + 1) Mod 256
        Signature = csp.SignData(PT, HA)
        Assert.That Sig, Iz.Not.EqualTo(Signature)
    Loop
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
            Case "ht": ht = ConvertHexToBytes(Trim$(Parts(1)))
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

