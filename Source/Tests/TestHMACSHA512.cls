VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHMACSHA512"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHMACSHA512
'
Option Explicit
Implements ICategorizable

Private csp As HMAC
Private FF As Long
Private PT() As Byte
Private ht() As Byte
Private KY() As Byte

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub


Public Sub Setup()
    Set csp = New HMACSHA512
    
    FF = FreeFile
    Open App.Path & "\Tests\Cryptography Files\HMACSHA512_Hashes.txt" For Input As #FF
End Sub

Public Sub Teardown()
    Close #FF
End Sub

Public Sub NewReturnsRandomKey()
    Dim Key() As Byte
    Key = csp.Key
    Assert.IsFalse CorArray.IsNull(Key)
    Assert.AreEqual 64, CorArray.GetLength(Key)
End Sub

Public Sub CannotUseNullKeyInConstructor()
    Dim b() As Byte
    On Error GoTo errTrap
    csp = Cor.NewHMACSHA512(b)
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
    Expected = Cor.NewBytes(&H44, &HD8, &H7, &H5B, &HFC, &HCB, &HCA, &H9, &HAF, &H2C, &H9A, &HC0, &H87, &HB9, &H3E, &HB, &HE, &HAB, &HC0, &H9F, &H9A, &H7E, &H32, &H5E, &H84, &H1D, &HCD, &H20, &H57, &HC6, &H3C, &HF, &HA0, &H21, &H48, &H38, &H6C, &H63, &H3B, &H14, &HD6, &H8E, &H34, &HFE, &H42, &HD9, &H8D, &H7, &HDA, &HD1, &HB4, &H47, &H4D, &H72, &HB8, &HB, &H55, &H5A, &H4, &H6A, &HC, &HB0, &HD5, &H1A)
    
    Assert.AreEqual Expected, csp.ComputeHash(Cor.NewBytes())
End Sub

Public Sub Test1ByteArray()
    csp.Key = Cor.NewBytes(0)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&HD7, &H43, &HBA, &H45, &HE9, &HF0, &H36, &H28, &HF8, &H30, &H5A, &H92, &H36, &HCD, &H1, &H82, &HA4, &H10, &HB2, &HD6, &H7D, &H70, &HC1, &HFF, &HB3, &H2C, &H6E, &H79, &H54, &H1A, &H34, &H24, &HA2, &H60, &H42, &H4, &H44, &HCA, &HDB, &HF, &H43, &H1A, &HAA, &HD5, &H42, &H5C, &HE3, &H13, &HCB, &HEA, &H9F, &H5E, &HD9, &HDA, &HE3, &HFD, &H88, &H33, &HDF, &HA6, &H75, &HD7, &H66, &H21)
    
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
    hash = Cor.NewBytes(&HAB, &H28, &HA5, &H8C, &H4C, &H77, &HDA, &H2A, &H91, &H42, &HCA, &H5A, &H5B, &H47, &HA7, &H2B, &H3E, &H54, &HFC, &H7D, &HFF, &H97, &H70, &HFA, &H59, &H1B, &HD, &H55, &HDC, &H28, &HCB, &HE, &HA4, &HCF, &H45, &HA3, &H7B, &HDE, &H4D, &H3B, &H2A, &H34, &HCE, &HBD, &H52, &H3, &HE, &H46, &HAB, &HC2, &H23, &HD2, &H44, &H40, &H87, &H17, &H9C, &H16, &H71, &HE5, &H6B, &H46, &HF0, &HF3)
    
    csp.Key = Cor.NewBytes(&H57, &H69, &H94, &H2C, &HA2, &H9C, &H8F, &H3D, &H16, &H2B, &H93, &H4E, &H59, &HA9, &H61, &HB0, &H83, &HBD, &HCE, &HBA, &HA2, &HF8, &H83, &HF1, &H46, &H9B, &HBC, &H22, &HAD, &H70, &H54, &HC6, &H72, &H6A, &HD7, &H3D, &H1B, &H77, &HCB, &H65, &H95, &HDA, &H6C, &H5F, &HAB, &H18, &HC0, &H2E, &H49, &HB6, &HDC, &HE7, &H69, &H12, &HEE, &HA8, &H1, &H47, &H54, &HF2, &HC, &HC2, &HD0, &H5)
    
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


