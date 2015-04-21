VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHMACSHA384"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHMACSHA384
'
Option Explicit
Private csp As HMAC
Private FF As Long
Private PT() As Byte
Private ht() As Byte
Private KY() As Byte


Public Sub Setup()
    Set csp = New HMACSHA384
    
    FF = FreeFile
    Open App.Path & "\Tests\Cryptography Files\HMACSHA384_Hashes.txt" For Input As #FF
End Sub

Public Sub TearDown()
    Close #FF
End Sub

Public Sub NewReturnsRandomKey()
    Dim Key() As Byte
    Key = csp.Key
    Assert.IsFalse cArray.IsNull(Key)
    Assert.AreEqual 64, cArray.GetLength(Key)
End Sub

Public Sub CannotUseNullKeyInConstructor()
    Dim b() As Byte
    On Error GoTo errTrap
    csp = Cor.NewHMACSHA384(b)
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is ArgumentNullException, "Wrong exception type caught."
End Sub

Public Sub TestEmptyArray()
    csp.Key = Cor.NewBytes(0)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&HEA, &H97, &HC7, &HFB, &H43, &H9E, &H3E, &H5F, &HC1, &H2C, &H61, &HF2, &HD, &HEF, &HA2, &H56, &H87, &HA9, &H8, &H56, &HFC, &H87, &H95, &HD3, &HE, &H13, &H5C, &HF6, &H9A, &H31, &H6C, &H3C, &H95, &H22, &H6B, &HC3, &H94, &HF5, &H2B, &H86, &H0, &H10, &HD7, &H45, &HAF, &HFD, &H67, &HF8)
    
    Assert.AreEqual Expected, csp.ComputeHash(Cor.NewBytes())
End Sub

Public Sub Test1ByteArray()
    csp.Key = Cor.NewBytes(0)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&HED, &H54, &H8F, &HD1, &HBE, &H20, &H4D, &HF8, &H58, &H69, &HC, &H61, &H75, &H66, &H73, &H88, &H7B, &H31, &H7B, &H6, &HB2, &H66, &H76, &H85, &H5, &H1B, &H3E, &H4C, &HCF, &HED, &H63, &HC, &HE9, &H3F, &H2F, &HCD, &HC1, &HE, &H5E, &H15, &H4D, &H5E, &H9F, &HA3, &H28, &H70, &H54, &H87)
    
    Assert.AreEqual Expected, csp.ComputeHash(Cor.NewBytes(1))
End Sub

Public Sub TestComputeHash()
    Do While GetNextEntry
        csp.Key = KY
        
        Assert.AreEqual ht, csp.ComputeHash(PT)
    Loop
End Sub

Public Sub TestTransformFinalBlock()
    Do While GetNextEntry
        csp.Key = KY
        
        Dim b() As Byte
        b = csp.TransformFinalBlock(PT, 0, UBound(PT) + 1)
        Assert.AreEqual PT, b
        Assert.AreEqual ht, csp.hash
    Loop
End Sub

Public Sub TestTransformBlock1ByteAtATime()
    Dim c As Long
    
    Do While GetNextEntry
        Dim i As Long
        Dim b() As Byte
        ReDim b(0)
        
        c = c + 1
        
        csp.Key = KY
        
        For i = 0 To UBound(PT)
            Assert.AreEqual 1, csp.TransformBlock(PT, i, 1, b, 0), "Entry: " & c
            Assert.AreEqual PT(i), b(0), "TransformBlock Entry: " & c
        Next i
        csp.TransformFinalBlock PT, UBound(PT) + 1, 0
        Assert.AreEqual ht, csp.hash, "TransformFinalBlock Entry: " & c
    Loop
End Sub

Public Sub TestTransformBlock5BytesAtATime()
    Do While GetNextEntry
        Dim c As Long
        Dim i As Long
        Dim b() As Byte
        ReDim b(4)
        
        c = c + 1
        
        csp.Key = KY
        
        i = 0
        Do While i <= (UBound(PT) - 5)
            Assert.AreEqual 5, csp.TransformBlock(PT, i, 5, b, 0)
            
            Dim j As Long
            For j = 0 To 4
                Assert.AreEqual PT(i + j), b(j), "TransformBlock: Entry " & c
            Next j
            
            i = i + 5
        Loop
        csp.TransformFinalBlock PT, i, UBound(PT) - i + 1
        Assert.AreEqual ht, csp.hash, "Hash: Entry " & c
    Loop
End Sub

Public Sub HashLargeFileStream()
    Dim hash() As Byte
    hash = Cor.NewBytes(&H1A, &H35, &HE, &HBA, &H8D, &H19, &H12, &H18, &HCA, &H55, &HCB, &HE5, &H50, &H79, &H1E, &HC9, &HF, &HB5, &H99, &H76, &H6D, &H11, &HE3, &HBC, &H4E, &HFD, &H7B, &H88, &HC6, &H5C, &H8D, &H1E, &HF2, &H96, &HF7, &HEC, &H3C, &HEC, &HD2, &HA7, &HF, &HC1, &H4C, &HC, &HC7, &H4, &H48, &H5C)
    
    csp.Key = Cor.NewBytes(&HC4, &HF1, &HE1, &HF7, &HC, &H7B, &H53, &H32, &HCA, &HD7, &HAE, &H8A, &H1, &H60, &HB3, &H54, &HB1, &HB7, &H56, &H10, &H2D, &HC9, &HBF, &HCA, &HAE, &HD4, &H2B, &HF0, &HE5, &H97, &H45, &HCF, &HA6, &H3C, &HF0, &H25, &H11, &H98, &HE5, &H3A, &H5, &H22, &H67, &H7C, &HF1, &HCA, &HBC, &H36, &H13, &H5, &H29, &HD0, &H74, &HE3, &H8E, &H29, &H17, &HAE, &HDF, &HDC, &H2, &H27, &H3A, &H15)
    
    Assert.AreEqual hash, csp.ComputeHash(File.OpenFile(App.Path & "\Tests\Cryptography Files\RC2_Entries.txt", OpenExisting))
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
            Case "ky": KY = ConvertHexToBytes(Trim$(Parts(1)))
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


