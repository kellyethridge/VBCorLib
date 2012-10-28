VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSHA512Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSHA512Managed
'
Option Explicit

Private PT() As Byte
Private ht() As Byte
Private csp As SHA512Managed
Private FF As Long

Public Sub Setup()
    Set csp = New SHA512Managed
    
    FF = FreeFile
    Open App.Path & "\Cryptography Files\SHA512_Hashes.txt" For Input As #FF
End Sub

Public Sub TearDown()
    Close #FF
End Sub


Public Sub TestPropertyValues()
    Assert.AreEqual 512, csp.HashSize
    Assert.AreEqual 1, csp.InputBlockSize
    Assert.AreEqual 1, csp.OutputBlockSize
    Assert.IsTrue csp.CanReuseTransform
    Assert.IsTrue csp.CanTransformMultipleBlocks
End Sub


Public Sub TestEmptyArray()
    Dim b() As Byte
    b = csp.ComputeHash(Cor.NewBytes())
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&HCF, &H83, &HE1, &H35, &H7E, &HEF, &HB8, &HBD, &HF1, &H54, &H28, &H50, &HD6, &H6D, &H80, &H7, &HD6, &H20, &HE4, &H5, &HB, &H57, &H15, &HDC, &H83, &HF4, &HA9, &H21, &HD3, &H6C, &HE9, &HCE, &H47, &HD0, &HD1, &H3C, &H5D, &H85, &HF2, &HB0, &HFF, &H83, &H18, &HD2, &H87, &H7E, &HEC, &H2F, &H63, &HB9, &H31, &HBD, &H47, &H41, &H7A, &H81, &HA5, &H38, &H32, &H7A, &HF9, &H27, &HDA, &H3E)
    
    Assert.AreEqual Expected, b
End Sub

Public Sub TestOneByte()
    Dim b() As Byte
    b = csp.ComputeHash(Cor.NewBytes(0))
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&HB8, &H24, &H4D, &H2, &H89, &H81, &HD6, &H93, &HAF, &H7B, &H45, &H6A, &HF8, &HEF, &HA4, &HCA, &HD6, &H3D, &H28, &H2E, &H19, &HFF, &H14, &H94, &H2C, &H24, &H6E, &H50, &HD9, &H35, &H1D, &H22, &H70, &H4A, &H80, &H2A, &H71, &HC3, &H58, &HB, &H63, &H70, &HDE, &H4C, &HEB, &H29, &H3C, &H32, &H4A, &H84, &H23, &H34, &H25, &H57, &HD4, &HE5, &HC3, &H84, &H38, &HF0, &HE3, &H69, &H10, &HEE)
    
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
    Assert.IsTrue catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is InvalidOperationException, "Wrong exception type."
End Sub

Public Sub HashLargeFileStream()
    Dim fs As FileStream
    Set fs = Cor.NewFileStream(App.Path & "\Cryptography Files\RC2_Entries.txt", OpenExisting)
    
    Dim hash() As Byte
    hash = Cor.NewBytes(&H19, &HDF, &HB7, &HEE, &H11, &HB0, &HFC, &H65, &H4D, &H56, &H4E, &H5, &H25, &HE, &H85, &HA, &H6F, &H3D, &H7F, &H72, &H6, &HD8, &HB1, &HF6, &H67, &HF, &H21, &HDA, &H84, &HF7, &H32, &HF, &HC7, &HB, &H29, &HDF, &H9C, &HC3, &H5D, &H16, &HF8, &HAF, &H6B, &H52, &H22, &HF2, &H18, &HA9, &H58, &HC9, &HD7, &HC6, &H19, &H26, &H70, &H70, &H73, &HF2, &HC3, &H72, &H18, &HCC, &H32, &HCB)
    
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

