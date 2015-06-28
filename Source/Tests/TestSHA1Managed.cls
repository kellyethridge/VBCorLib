VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSHA1Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSHA1Managed
'
Option Explicit
Implements ICategorizable

Private PT() As Byte
Private ht() As Byte
Private csp As SHA1Managed
Private FF As Long

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Setup()
    Set csp = New SHA1Managed
    
    FF = FreeFile
    Open App.Path & "\Tests\Cryptography Files\SHA1_Hashes.txt" For Input As #FF
End Sub

Public Sub Teardown()
    Close #FF
End Sub


Public Sub TestPropertyValues()
    Assert.AreEqual 160, csp.HashSize
    Assert.AreEqual 1, csp.InputBlockSize
    Assert.AreEqual 1, csp.OutputBlockSize
    Assert.IsTrue csp.CanReuseTransform
    Assert.IsTrue csp.CanTransformMultipleBlocks
End Sub

Public Sub TestNoData()
    PT = Cor.NewBytes
    ht = csp.ComputeHash(PT)
    Assert.AreEqual ConvertHexToBytes("da39a3ee5e6b4b0d3255bfef95601890afd80709"), ht
End Sub

Public Sub Test1Byte()
    ReDim PT(0)
    ht = csp.ComputeHash(PT)
    Assert.AreEqual 20, CorArray.Length(ht)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&H5B, &HA9, &H3C, &H9D, &HB0, &HCF, &HF9, &H3F, &H52, &HB5, &H21, &HD7, &H42, &HE, &H43, &HF6, &HED, &HA2, &H78, &H4F)
    
    Assert.AreEqual Expected, ht
End Sub

Public Sub TestSentence()
    Dim b() As Byte
    b = Encoding.ASCII.GetBytes("The quick brown fox jumps over the lazy dog")
    ht = csp.ComputeHash(b)
    Assert.AreEqual ConvertHexToBytes("2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"), ht
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
    hash = Cor.NewBytes(&H37, &HEB, &H94, &HAB, &H1A, &H4C, &HC7, &HC4, &H50, &H3A, &H4D, &H70, &HF5, &HC2, &HA3, &H41, &H4D, &HB2, &H97, &H4A)
    
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


