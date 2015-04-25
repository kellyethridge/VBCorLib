VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHMACMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHMACMD5
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
    Set csp = New HMACMD5
    
    FF = FreeFile
    Open App.Path & "\Tests\Cryptography Files\HMACMD5_Hashes.txt" For Input As #FF
End Sub

Public Sub Teardown()
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
    Set csp = Cor.NewHMACMD5(b)
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
    Expected = Cor.NewBytes(&H74, &HE6, &HF7, &H29, &H8A, &H9C, &H2D, &H16, &H89, &H35, &HF5, &H8C, &H0, &H1B, &HAD, &H88)
    
    Assert.AreEqual Expected, csp.ComputeHash(Cor.NewBytes())
End Sub

Public Sub Test1ByteArray()
    csp.Key = Cor.NewBytes(0)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&HB4, &H11, &H53, &HDC, &H48, &H51, &HA, &HEE, &HE0, &H89, &HEE, &H50, &HD9, &H57, &H44, &H34)
    
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
    hash = Cor.NewBytes(&HAB, &HDB, &H30, &H5C, &HA8, &HF, &HC9, &HD3, &H46, &H72, &H14, &H5F, &H5F, &H54, &HF, &H59)
    
    csp.Key = Cor.NewBytes(&H69, &HD4, &H77, &H7A, &H99, &H54, &H73, &HE8, &H2B, &HDA, &HC8, &H3F, &H88, &HBE, &H40, &HB0, &HC7, &H45, &H3B, &HD7, &H2C, &HBA, &HA3, &H67, &H5C, &H7, &HE1, &HB4, &HF8, &H52, &H64, &HA6, &HB9, &H56, &HFD, &HE2, &H90, &HDA, &HAD, &HA3, &HB3, &H4E, &HE6, &H75, &HF1, &H7A, &HD1, &H6B, &H86, &H4D, &HB, &H6F, &HB6, &HFA, &HE5, &H55, &H0, &HBA, &H99, &HD3, &H3E, &HF1, &H7B, &H6)
    
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


