VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTripleDESDecryption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTripleDESDecryption
'
Option Explicit
Implements ICategorizable

Private csp As TripleDESCryptoServiceProvider
Private FF As Long
Private PT() As Byte
Private CT() As Byte

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub


Public Sub Setup()
    Set csp = New TripleDESCryptoServiceProvider
    FF = FreeFile
    Open App.Path & "\Cryptography Files\TripleDES_Entries.txt" For Input As #FF
End Sub

Public Sub Teardown()
    Close #FF
    Erase PT
    Erase CT
End Sub

Public Sub TestTransformFinalBlock()
    Dim c As Long
    
    Do While GetNextEntry()
        PadWithZeros PT
        
        c = c + 1
        
        Dim t As ICryptoTransform
        Set t = csp.CreateDecryptor
        Dim NewPT() As Byte
        On Error GoTo errTrap
        NewPT = t.TransformFinalBlock(CT, 0, UBound(CT) + 1)
        On Error GoTo 0
        
        Assert.AreEqual PT, NewPT, "Entry: " & c
    Loop
    
    WriteLine "TripleDES TransformFinalBlock Decryption: " & c & " combinations tested."
    Exit Sub
    
errTrap:
    Assert.Fail "Entry " & c & ": " & Err.Description
End Sub

Public Sub TestTransformBlock()
    Dim c As Long
    
    Do While GetNextEntry()
        PadWithZeros PT
        
        c = c + 1
        
        Dim t As ICryptoTransform
        Set t = csp.CreateDecryptor
        Dim NewPT() As Byte
        Dim AmountTransformed As Long
        
        ReDim NewPT(UBound(CT))
        AmountTransformed = t.TransformBlock(CT, 0, UBound(CT) + 1, NewPT, 0)
        
        If (csp.Padding = PaddingMode.Zeros) Or (csp.Padding = PaddingMode.None) Then
            Assert.AreEqual UBound(CT) + 1, AmountTransformed
        Else
            Assert.AreEqual (UBound(CT) + 1) - (csp.BlockSize \ 8), AmountTransformed, "Wrong amount transformed."
            If AmountTransformed = 0 Then
                NewPT = cArray.CreateInstance(vbByte)
            Else
                ReDim Preserve NewPT(0 To AmountTransformed - 1)
            End If
        End If
        
        ' Get the final block
        Dim Final() As Byte
        Final = t.TransformFinalBlock(CT, UBound(CT) + 1, 0)
        
        ' Append the final block to the completed decrypted array.
        ReDim Preserve NewPT(0 To UBound(NewPT) + UBound(Final) + 1)
        Buffer.BlockCopy Final, 0, NewPT, AmountTransformed, UBound(Final) + 1
        
        Assert.AreEqual PT, NewPT, "Entry " & c & ": newPT does not equal PT."
    Loop
End Sub

Public Sub TestTransformBlockMultipleCalls()
    Dim c As Long
    
    Do While GetNextEntry()
        PadWithZeros PT
        
        c = c + 1
        
        Dim t As ICryptoTransform
        Set t = csp.CreateDecryptor
        
        Dim NewPT() As Byte
        ReDim NewPT(UBound(CT))
        
        Dim CTIndex As Long
        Dim PTIndex As Long
        
        CTIndex = 0
        PTIndex = 0
        
        Do While CTIndex < UBound(CT)
            Dim AmountTransformed As Long
            AmountTransformed = t.TransformBlock(CT, CTIndex, t.InputBlockSize, NewPT, PTIndex)
            CTIndex = CTIndex + t.InputBlockSize
            PTIndex = PTIndex + AmountTransformed
        Loop
        
        Dim FinalBlock() As Byte
        FinalBlock = t.TransformFinalBlock(CT, CTIndex, UBound(CT) - CTIndex + 1)
        
        Buffer.BlockCopy FinalBlock, 0, NewPT, PTIndex, cArray.GetLength(FinalBlock)
        
        If (csp.Padding <> PaddingMode.None) And (csp.Padding <> PaddingMode.Zeros) Then
            ReDim Preserve NewPT(0 To UBound(NewPT) - (t.OutputBlockSize - cArray.GetLength(FinalBlock)))
        End If
        
        Assert.AreEqual PT, NewPT, "Plain text entry " & c & " is incorrect."
    Loop
End Sub

Public Sub ResetForNewText()
    Do
        GetNextEntry
    Loop While cArray.GetLength(CT) < 16
    
    Dim NewPT() As Byte
    ReDim NewPT(15)
    ReDim Preserve PT(7)
    
    Dim t As CryptoAPITransform
    Set t = csp.CreateDecryptor
    
    Dim i As Long
    For i = 1 To 5
        ReDim NewPT(15)
        Assert.AreEqual 8, t.TransformBlock(CT, 0, 16, NewPT, 0)
        ReDim Preserve NewPT(7)
        Assert.AreEqual PT, NewPT
        t.Reset
    Next i
End Sub



Private Sub PadWithZeros(ByRef PT() As Byte)
    ' We pad the plain text with zeros if necessary because the decrypted text
    ' will have the zero padding and we don't know what to remove.
    If csp.Padding = PaddingMode.Zeros Then
        If ((UBound(PT) + 1) Mod csp.BlockSize \ 8) <> 0 Then
            ReDim Preserve PT(0 To UBound(PT) + (csp.BlockSize \ 8 - ((UBound(PT) + 1) Mod csp.BlockSize \ 8)))
        End If
    End If
End Sub

Private Function GetNextEntry() As Boolean
    If EOF(FF) Then Exit Function
    
    Dim line As String
    Line Input #FF, line
    Do While Len(line) > 0
        Dim Parts() As String
        Parts = Split(line, "=")
        Select Case LCase$(Trim$(Parts(0)))
            Case "md": csp.Mode = CInt(Parts(1))
            Case "pd": csp.Padding = CInt(Parts(1))
            Case "bz": csp.BlockSize = CInt(Parts(1))
            Case "kz": csp.KeySize = CInt(Parts(1))
            Case "ky": csp.Key = ConvertHexToBytes(Trim$(Parts(1)))
            Case "iv": csp.IV = ConvertHexToBytes(Trim$(Parts(1)))
            Case "pt": PT = ConvertHexToBytes(Trim$(Parts(1)))
            Case "ct": CT = ConvertHexToBytes(Trim$(Parts(1)))
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


