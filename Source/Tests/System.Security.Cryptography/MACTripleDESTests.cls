VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACTripleDESTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MACTripleDESTests
'
Option Explicit
Implements ICategorizable

'Private csp As MACTripleDes
'Private PT() As Byte
'Private ht() As Byte
'Private FF As Long

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

'Public Sub Setup()
'    Set Csp = New MACTripleDes
'
'    FF = FreeFile
'    Open App.Path & "\Cryptography Files\MACTripleDES_Hashes.txt" For Input As #FF
'End Sub
'
'Public Sub Teardown()
'    Close #FF
'End Sub

'Public Sub TestEmptyArray()
'    Dim i As Long
'    Dim Key() As Byte
'    ReDim Key(0 To 23)
'    For i = 0 To 23
'        Key(i) = i
'    Next i
'
'    csp.Key = Key
'    csp.Padding = PKCS7
'
'    Dim Expected() As Byte
'    Expected = Cor.NewBytes(&HA3, &HCF, &H6E, &HC8, &H8B, &HD9, &H7D, &H73)
'
'    Assert.AreEqual Expected, csp.ComputeHash(Cor.NewBytes())
'End Sub
'
'Public Sub TestComputeHashesFromBytes()
'    Dim c As Long
'
'    Do While GetNextEntry
'        c = c + 1
'
'        Dim b() As Byte
'        b = csp.ComputeHash(PT)
'        Assert.IsFalse CorArray.IsNull(b), "b should not be null."
'        Assert.AreEqual ht, b, "Entry: " & c
'    Loop
'End Sub
'
'Public Sub TestTransformFinalBlock()
'    Do While GetNextEntry
'
'        Dim b() As Byte
'        b = csp.TransformFinalBlock(PT, 0, UBound(PT) + 1)
'        Assert.AreEqual PT, b
'        Assert.AreEqual ht, csp.Hash
'    Loop
'End Sub
'
'Public Sub TestTransformBlock1ByteAtATime()
'    Dim c As Long
'
'    Do While GetNextEntry
'        Dim i As Long
'        Dim b() As Byte
'        ReDim b(0)
'
'        c = c + 1
'
'        For i = 0 To UBound(PT)
'            Assert.AreEqual 1, csp.TransformBlock(PT, i, 1, b, 0), "Entry: " & c
'            Assert.AreEqual PT(i), b(0), "Entry: " & c
'        Next i
'        csp.TransformFinalBlock PT, UBound(PT) + 1, 0
'        Assert.AreEqual ht, csp.Hash, "Entry: " & c
'    Loop
'End Sub
'
'Public Sub TestTransformBlock5BytesAtATime()
'    Do While GetNextEntry
'        Dim c As Long
'        Dim i As Long
'        Dim b() As Byte
'        ReDim b(4)
'
'        c = c + 1
'
'        i = 0
'        Do While i <= (UBound(PT) - 5)
'            Assert.AreEqual 5, csp.TransformBlock(PT, i, 5, b, 0)
'
'            Dim j As Long
'            For j = 0 To 4
'                Assert.AreEqual PT(i + j), b(j), "TransformBlock: Entry " & c
'            Next j
'
'            i = i + 5
'        Loop
'        csp.TransformFinalBlock PT, i, UBound(PT) - i + 1
'        Assert.AreEqual ht, csp.Hash, "Hash: Entry " & c
'    Loop
'End Sub
'
'Public Sub HashLargeFileStream()
'    Dim fs As FileStream
'    Set fs = Cor.NewFileStream(App.Path & "\Cryptography Files\RC2_Entries.txt", OpenExisting)
'
'    Dim Hash() As Byte
'    Hash = Cor.NewBytes(&H2B, &H6F, &H35, &HC8, &H18, &H3E, &H37, &H1D)
'
'    csp.Key = Cor.NewBytes(&H4B, &H60, &H72, &HD3, &HD8, &H76, &HAD, &H6C, &H33, &HC1, &H9D, &H53, &HDD, &HEA, &H31, &HD9, &H25, &HC0, &HFE, &H9C, &HCB, &H4D, &H43, &H1D)
'
'    Assert.AreEqual Hash, csp.ComputeHash(fs)
'End Sub
'
'
'
'
'
'
'
'Private Function GetNextEntry() As Boolean
'    If EOF(FF) Then Exit Function
'
'    Dim line As String
'    Line Input #FF, line
'    Do While Len(line) > 0
'        Dim Parts() As String
'        Parts = Split(line, "=")
'        Select Case LCase$(Trim$(Parts(0)))
'            Case "pt": PT = ConvertHexToBytes(Trim$(Parts(1)))
'            Case "ht": ht = ConvertHexToBytes(Trim$(Parts(1)))
'            Case "ky": csp.Key = ConvertHexToBytes(Trim$(Parts(1)))
'            Case "pd": csp.Padding = CInt(Parts(1))
'        End Select
'
'        Line Input #FF, line
'    Loop
'    GetNextEntry = True
'End Function
'
'Private Function ConvertHexToBytes(ByVal s As String) As Byte()
'    Dim ret() As Byte
'    ReDim ret(0 To Len(s) \ 2 - 1)
'
'    Dim i As Long
'    For i = 0 To UBound(ret)
'        ret(i) = CByte("&h" & Mid$(s, (i * 2) + 1, 2))
'    Next i
'
'    ConvertHexToBytes = ret
'End Function
'
'Private Function ConvertBytesToHex(ByRef Bytes() As Byte) As String
'    Dim i As Long
'    Dim sb As New StringBuilder
'
'    For i = 0 To UBound(Bytes)
'        sb.AppendFormat "{0:X2}", Bytes(i)
'    Next i
'
'    ConvertBytesToHex = sb.ToString
'End Function
'
'
'
