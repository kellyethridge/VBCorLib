VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRIPEMD160Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRIPEMD160Managed
'
Option Explicit
Private PT() As Byte
Private ht() As Byte
Private csp As RIPEMD160
Private FF As Long

Public Sub Setup()
    Set csp = New RIPEMD160Managed
    
    FF = FreeFile
    Open App.Path & "\Cryptography Files\RIPEMD160_Hashes.txt" For Input As #FF
End Sub

Public Sub TearDown()
    Close #FF
End Sub


Public Sub TestPropertyValues()
    Assert.AreEqual 160, csp.HashSize
    Assert.AreEqual 1, csp.InputBlockSize
    Assert.AreEqual 1, csp.OutputBlockSize
    Assert.IsTrue csp.CanReuseTransform
    Assert.IsTrue csp.CanTransformMultipleBlocks
End Sub

Public Sub TestEmptyArray()
    Dim csp As New RIPEMD160Managed
    Dim ht() As Byte
    ht = csp.ComputeHash(Cor.NewBytes())
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&H9C, &H11, &H85, &HA5, &HC5, &HE9, &HFC, &H54, &H61, &H28, &H8, &H97, &H7E, &HE8, &HF5, &H48, &HB2, &H25, &H8D, &H31)
    
    Assert.AreEqual Expected, ht
End Sub

Public Sub TestOneByte()
    Dim b() As Byte
    b = csp.ComputeHash(Cor.NewBytes(1))
    
    Dim ht() As Byte
    ht = Cor.NewBytes(&HF2, &H91, &HBA, &H50, &H15, &HDF, &H34, &H8C, &H80, &H85, &H3F, &HA5, &HBB, &HF, &H79, &H46, &HF5, &HC9, &HE1, &HB3)
    PrintBytes b
    
    Assert.AreEqual ht, b
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
        Dim c As Long
        c = c + 1
        
        b = csp.TransformFinalBlock(PT, 0, UBound(PT) + 1)
        Assert.AreEqual PT, b
        Assert.AreEqual ht, csp.Hash, "Entry " & c
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
        Assert.AreEqual ht, csp.Hash, "Entry: " & c
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
        Assert.AreEqual ht, csp.Hash, "Hash: Entry " & c
    Loop
End Sub

Public Sub GetHashBeforeFinishedShouldFail()
    On Error GoTo errTrap
    Dim b() As Byte
    b = csp.Hash
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is InvalidOperationException, "Wrong exception type."
End Sub

Public Sub HashLargeFileStream()
    Dim fs As FileStream
    Set fs = Cor.NewFileStream(App.Path & "\Cryptography Files\RC2_Entries.txt", OpenExisting)
    
    Dim Hash() As Byte
    Hash = Cor.NewBytes(&HC0, &H63, &H27, &HD3, &HC8, &H54, &HC7, &H2A, &HF4, &HF7, &HD6, &HB8, &HC5, &H0, &H60, &HB1, &H7, &HAD, &H9A, &H39)
    
    Assert.AreEqual Hash, csp.ComputeHash(fs)
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



Private Sub PrintBytes(ByRef Bytes() As Byte)
    Dim i As Long
    For i = LBound(Bytes) To UBound(Bytes)
        Debug.Print cString.Format("{0:X2} ", Bytes(i));
    Next i
    Debug.Print
End Sub

