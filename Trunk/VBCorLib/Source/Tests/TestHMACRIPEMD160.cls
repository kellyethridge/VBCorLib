VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHMACRIPEMD160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHMACRIPEMD160
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
    Set csp = New HMACRIPEMD160
    
    FF = FreeFile
    Open App.Path & "\Tests\Cryptography Files\HMACRIPEMD160_Hashes.txt" For Input As #FF
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
    csp = Cor.NewHMACRIPEMD160(b)
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
    Expected = Cor.NewBytes(&H44, &HD8, &H6B, &H65, &H8A, &H3E, &H7C, &HBC, &H1A, &H20, &H10, &H84, &H8B, &H53, &HE3, &H5C, &H91, &H77, &H20, &HCA)
    
    Assert.AreEqual Expected, csp.ComputeHash(Cor.NewBytes())
End Sub

Public Sub Test1ByteArray()
    csp.Key = Cor.NewBytes(0)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&HEB, &HC5, &H1D, &H89, &HFF, &H81, &H55, &H2A, &HB3, &H88, &H81, &HE1, &HA1, &H2B, &H16, &H17, &HDD, &H83, &HB7, &H9F)
    
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
    hash = Cor.NewBytes(&H39, &HCA, &H8A, &H2D, &H93, &HF9, &H5, &H6C, &HE1, &H22, &H63, &H0, &H2F, &HED, &HC, &H33, &H93, &H1E, &H3B, &H24)
    
    csp.Key = Cor.NewBytes(&H5A, &H41, &HDC, &HDA, &H95, &H50, &HEA, &H8E, &H7A, &HFF, &HF2, &H34, &H5, &H0, &H94, &H95, &H27, &HCD, &H1E, &H89, &H36, &H26, &H4F, &H84, &HB9, &HB1, &HBC, &H94, &HC5, &H59, &HB3, &H82, &HD2, &H94, &H87, &H3, &HE9, &H2F, &HA0, &H57, &H47, &HCD, &H69, &HE9, &HAF, &H6A, &H71, &HA9, &HD6, &H7F, &H90, &HBF, &HA2, &H87, &H25, &H87, &H89, &HC8, &H24, &H5F, &HA6, &H98, &HA7, &H21)
    
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


