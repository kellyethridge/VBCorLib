VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHMACSHA256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHMACSHA256
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
    Set csp = New HMACSHA256
    
    FF = FreeFile
    Open App.Path & "\Tests\Cryptography Files\HMACSHA256_Hashes.txt" For Input As #FF
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
    csp = Cor.NewHMACSHA256(b)
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
    Expected = Cor.NewBytes(&HB6, &H13, &H67, &H9A, &H8, &H14, &HD9, &HEC, &H77, &H2F, &H95, &HD7, &H78, &HC3, &H5F, &HC5, &HFF, &H16, &H97, &HC4, &H93, &H71, &H56, &H53, &HC6, &HC7, &H12, &H14, &H42, &H92, &HC5, &HAD)
    
    Assert.AreEqual Expected, csp.ComputeHash(Cor.NewBytes())
End Sub

Public Sub Test1ByteArray()
    csp.Key = Cor.NewBytes(0)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&H3D, &H7A, &HFB, &H66, &H31, &H24, &HEC, &HBF, &H2C, &H95, &H3F, &H86, &H3D, &H4F, &HC8, &H79, &H6E, &HEB, &H2D, &H37, &H2B, &H64, &HAA, &HD5, &H86, &H97, &HEC, &H52, &H64, &H64, &H9C, &HDB)
    
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
    hash = Cor.NewBytes(&H33, &H52, &H81, &HE, &HD1, &H4D, &H0, &H93, &HB0, &HC3, &HD1, &H4E, &H5D, &HEA, &H27, &H7F, &H73, &H10, &H58, &H83, &H47, &H34, &H85, &H65, &H55, &HF1, &HA7, &HF3, &HF6, &H6F, &H74, &H7B)
    
    csp.Key = Cor.NewBytes(&HEC, &HBD, &H11, &HEB, &H99, &HB4, &H28, &H27, &H5B, &H23, &H1A, &HC0, &HB4, &HF8, &H40, &H84, &H5, &H2D, &HC3, &HB6, &H8B, &HD9, &HB1, &HE1, &H5, &HE0, &HB2, &H8E, &HE3, &H5A, &H5D, &H42, &H63, &H85, &H73, &H3, &HDF, &H4E, &H20, &HDB, &HCC, &HE1, &H8B, &HBC, &HE7, &H1A, &H59, &H2D, &H1D, &HB, &HEE, &HBB, &HCB, &H6, &H57, &HC7, &H36, &H56, &HD8, &H9, &H5E, &H68, &HCD, &HCC)
    
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

