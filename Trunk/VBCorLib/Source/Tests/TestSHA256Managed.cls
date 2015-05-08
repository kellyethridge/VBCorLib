VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSHA256Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSHA256Managed
'
Option Explicit
Implements ICategorizable

Private PT() As Byte
Private ht() As Byte
Private csp As SHA256Managed
Private FF As Long

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Setup()
    Set csp = New SHA256Managed
    
    FF = FreeFile
    Open App.Path & "\Tests\Cryptography Files\SHA256_Hashes.txt" For Input As #FF
End Sub

Public Sub Teardown()
    Close #FF
End Sub


Public Sub TestPropertyValues()
    Assert.AreEqual 256, csp.HashSize
    Assert.AreEqual 1, csp.InputBlockSize
    Assert.AreEqual 1, csp.OutputBlockSize
    Assert.IsTrue csp.CanReuseTransform
    Assert.IsTrue csp.CanTransformMultipleBlocks
End Sub

Public Sub TestEmptyArray()
    Dim b() As Byte
    b = csp.ComputeHash(Cor.NewBytes())
    
    Dim ht() As Byte
    ht = Cor.NewBytes(&HE3, &HB0, &HC4, &H42, &H98, &HFC, &H1C, &H14, &H9A, &HFB, &HF4, &HC8, &H99, &H6F, &HB9, &H24, &H27, &HAE, &H41, &HE4, &H64, &H9B, &H93, &H4C, &HA4, &H95, &H99, &H1B, &H78, &H52, &HB8, &H55)
    
    Assert.AreEqual ht, b
End Sub

Public Sub TestOneByte()
    Dim b() As Byte
    b = csp.ComputeHash(Cor.NewBytes(1))
    
    Dim ht() As Byte
    ht = Cor.NewBytes(&H4B, &HF5, &H12, &H2F, &H34, &H45, &H54, &HC5, &H3B, &HDE, &H2E, &HBB, &H8C, &HD2, &HB7, &HE3, &HD1, &H60, &HA, &HD6, &H31, &HC3, &H85, &HA5, &HD7, &HCC, &HE2, &H3C, &H77, &H85, &H45, &H9A)
    
    Assert.AreEqual ht, b
End Sub

Public Sub TestComputeHashesFromBytes()
    Do While GetNextEntry
        Dim b() As Byte
        b = csp.ComputeHash(PT)
        Assert.IsFalse CorArray.IsNull(b), "b should not be null."
        Assert.AreEqual ht, b
    Loop
End Sub

Public Sub TestTransformFinalBlock()
    Do While GetNextEntry
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
        
        For i = 0 To UBound(PT)
            Assert.AreEqual 1, csp.TransformBlock(PT, i, 1, b, 0), "Entry: " & c
            Assert.AreEqual PT(i), b(0), "Entry: " & c
        Next i
        csp.TransformFinalBlock PT, UBound(PT) + 1, 0
        Assert.AreEqual ht, csp.hash, "Entry: " & c
    Loop
End Sub

Public Sub TestTransformBlock5BytesAtATime()
    Do While GetNextEntry
        Dim c As Long
        Dim i As Long
        Dim b() As Byte
        ReDim b(4)
        
        c = c + 1
        
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

Public Sub GetHashBeforeFinishedShouldFail()
    On Error GoTo errTrap
    Dim b() As Byte
    b = csp.hash
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is InvalidOperationException, "Wrong exception type."
End Sub

Public Sub HashLargeFileStream()
    Dim fs As FileStream
    Set fs = Cor.NewFileStream(App.Path & "\Tests\Cryptography Files\RC2_Entries.txt", OpenExisting)
    
    Dim hash() As Byte
    hash = Cor.NewBytes(&H44, &H41, &HB3, &H7B, &HE4, &HE, &HF3, &H8C, &HA3, &H7B, &H8F, &H24, &HC4, &HB5, &H9A, &HDF, &H5B, &HCE, &H4, &H37, &HB4, &H3B, &HEE, &H2D, &H7A, &H3, &H8E, &H4E, &H1C, &H2A, &H41, &HEC)
    
    Assert.AreEqual hash, csp.ComputeHash(fs)
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
