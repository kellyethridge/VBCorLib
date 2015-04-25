VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSHA384Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSHA384Managed
'
Option Explicit
Implements ICategorizable

Private PT() As Byte
Private ht() As Byte
Private csp As SHA384Managed
Private FF As Long

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub Setup()
    Set csp = New SHA384Managed
    
    FF = FreeFile
    Open App.Path & "\Tests\Cryptography Files\SHA384_Hashes.txt" For Input As #FF
End Sub

Public Sub Teardown()
    Close #FF
End Sub


Public Sub TestPropertyValues()
    Assert.AreEqual 384, csp.HashSize
    Assert.AreEqual 1, csp.InputBlockSize
    Assert.AreEqual 1, csp.OutputBlockSize
    Assert.IsTrue csp.CanReuseTransform
    Assert.IsTrue csp.CanTransformMultipleBlocks
End Sub


Public Sub TestEmptyArray()
    Dim b() As Byte
    b = csp.ComputeHash(Cor.NewBytes())
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&H38, &HB0, &H60, &HA7, &H51, &HAC, &H96, &H38, &H4C, &HD9, &H32, &H7E, &HB1, &HB1, &HE3, &H6A, &H21, &HFD, &HB7, &H11, &H14, &HBE, &H7, &H43, &H4C, &HC, &HC7, &HBF, &H63, &HF6, &HE1, &HDA, &H27, &H4E, &HDE, &HBF, &HE7, &H6F, &H65, &HFB, &HD5, &H1A, &HD2, &HF1, &H48, &H98, &HB9, &H5B)
    
    Assert.AreEqual Expected, b
End Sub

Public Sub TestOneByte()
    Dim b() As Byte
    b = csp.ComputeHash(Cor.NewBytes(0))
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&HBE, &HC0, &H21, &HB4, &HF3, &H68, &HE3, &H6, &H91, &H34, &HE0, &H12, &HC2, &HB4, &H30, &H70, &H83, &HD3, &HA9, &HBD, &HD2, &H6, &HE2, &H4E, &H5F, &HD, &H86, &HE1, &H3D, &H66, &H36, &H65, &H59, &H33, &HEC, &H2B, &H41, &H34, &H65, &H96, &H68, &H17, &HA9, &HC2, &H8, &HA1, &H17, &H17)
    
    Assert.AreEqual Expected, b
End Sub

Public Sub TestComputeHashesFromBytes()
    Do While GetNextEntry
        Dim b() As Byte
        b = csp.ComputeHash(PT)
        Assert.IsFalse cArray.IsNull(b), "b should not be null."
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
    hash = Cor.NewBytes(&HE2, &HFF, &HC7, &HEC, &HE5, &H8E, &H3A, &H67, &H7D, &HD4, &H58, &HCC, &HEB, &H32, &HEB, &HBB, &H38, &HAD, &HCD, &H2, &HD3, &HC5, &HC1, &HA0, &HA0, &HA3, &H45, &HE4, &HB7, &H9D, &H6D, &HD0, &HFF, &H1, &H1E, &H73, &H14, &HF8, &H98, &H9B, &H3D, &H4B, &H70, &HAE, &HC4, &HD0, &H5A, &H27)
    
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


