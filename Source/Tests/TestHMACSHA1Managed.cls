VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHMACSHA1Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHMACSHA1Managed
'
Option Explicit
Implements ICategorizable

Private Csp As HMAC
Private FF As Long
Private PT() As Byte
Private ht() As Byte
Private KY() As Byte

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub


'Public Sub Setup()
'    Set Csp = Cor.NewHMACSHA1(Cor.NewBytes(), True)
'
'    FF = FreeFile
'    Open App.Path & "\Tests\Cryptography Files\HMACSHA1_Hashes.txt" For Input As #FF
'End Sub
'
'Public Sub Teardown()
'    Close #FF
'End Sub

Public Sub TestEmptyArray()
    Csp.Key = Cor.NewBytes(0)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&HFB, &HDB, &H1D, &H1B, &H18, &HAA, &H6C, &H8, &H32, &H4B, &H7D, &H64, &HB7, &H1F, &HB7, &H63, &H70, &H69, &HE, &H1D)
    
    Assert.AreEqual Expected, Csp.ComputeHash(Cor.NewBytes())
End Sub

Public Sub Test1ByteArray()
    Csp.Key = Cor.NewBytes(0)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&H65, &H64, &HED, &H65, &H29, &H52, &H8D, &HF3, &HC4, &H2, &H78, &HB3, &HCC, &HD8, &H4C, &HC5, &H9B, &H6E, &H9C, &HCC)
    
    Assert.AreEqual Expected, Csp.ComputeHash(Cor.NewBytes(1))
End Sub

Public Sub TestComputeHash()
    Do While GetNextEntry
        Csp.Key = KY
        
        Assert.AreEqual ht, Csp.ComputeHash(PT)
    Loop
End Sub

Public Sub TestTransformFinalBlock()
    Do While GetNextEntry
        Csp.Key = KY
        
        Dim b() As Byte
        b = Csp.TransformFinalBlock(PT, 0, UBound(PT) + 1)
        Assert.AreEqual PT, b
        Assert.AreEqual ht, Csp.Hash
    Loop
End Sub

Public Sub TestTransformBlock1ByteAtATime()
    Dim c As Long
    
    Do While GetNextEntry
        Dim i As Long
        Dim b() As Byte
        ReDim b(0)
        
        c = c + 1
        
        Csp.Key = KY
        
        For i = 0 To UBound(PT)
            Assert.AreEqual 1, Csp.TransformBlock(PT, i, 1, b, 0), "Entry: " & c
            Assert.AreEqual PT(i), b(0), "TransformBlock Entry: " & c
        Next i
        Csp.TransformFinalBlock PT, UBound(PT) + 1, 0
        Assert.AreEqual ht, Csp.Hash, "TransformFinalBlock Entry: " & c
    Loop
End Sub

Public Sub TestTransformBlock5BytesAtATime()
    Do While GetNextEntry
        Dim c As Long
        Dim i As Long
        Dim b() As Byte
        ReDim b(4)
        
        c = c + 1
        
        Csp.Key = KY
        
        i = 0
        Do While i <= (UBound(PT) - 5)
            Assert.AreEqual 5, Csp.TransformBlock(PT, i, 5, b, 0)
            
            Dim j As Long
            For j = 0 To 4
                Assert.AreEqual PT(i + j), b(j), "TransformBlock: Entry " & c
            Next j
            
            i = i + 5
        Loop
        Csp.TransformFinalBlock PT, i, UBound(PT) - i + 1
        Assert.AreEqual ht, Csp.Hash, "Hash: Entry " & c
    Loop
End Sub

Public Sub HashLargeFileStream()
    Dim Hash() As Byte
    Hash = Cor.NewBytes(&H7, &HD2, &H6B, &H83, &H22, &HC6, &H6E, &HBC, &H3, &HF8, &HCA, &H55, &H34, &HB5, &H9D, &HA1, &H89, &HF6, &HD0, &H9A)
    
    Csp.Key = Cor.NewBytes(&H85, &HC1, &H5E, &H35, &H78, &H7C, &HA3, &HBE, &HE9, &H1E, &H99, &H14, &H6F, &HBE, &H78, &HCC, &HAF, &HD1, &H22, &HD3, &H29, &H8F, &HDC, &HEF, &H68, &HBC, &H5A, &HB8, &H5A, &HD8, &HCB, &HB3, &HD6, &HCC, &HD3, &HCD, &H7, &HDF, &H9D, &HD9, &H3D, &H8C, &H64, &HDD, &H7B, &H29, &H2A, &H4, &H3D, &HA4, &H36, &HA8, &H31, &HAC, &HCC, &HD8, &H68, &H43, &HB4, &H9D, &HDE, &HB1, &HA9, &HE5)
    
    Assert.AreEqual Hash, Csp.ComputeHash(File.OpenFile(App.Path & "\Tests\Cryptography Files\RC2_Entries.txt", OpenExisting))
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

