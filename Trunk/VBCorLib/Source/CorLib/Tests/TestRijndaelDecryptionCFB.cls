VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRijndaelDecryptionCFB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRijndaelDecryptionCFB
'
Option Explicit

Private csp As RijndaelManaged
Private FF As Long
Private PT() As Byte
Private CT() As Byte


Public Sub Setup()
    Set csp = New RijndaelManaged
    FF = FreeFile
    Open App.Path & "\Cryptography Files\Rijndael_CFB_Entries.txt" For Input As #FF
End Sub

Public Sub TearDown()
    Close #FF
    Erase PT
    Erase CT
End Sub

Public Sub TestTransformFinalBlock()
    Dim c As Long
    
    On Error GoTo errTrap
    Do While GetNextEntry()
        PadWithZeros PT
        
        c = c + 1
        
        Dim t As ICryptoTransform
        Set t = csp.CreateDecryptor
        Dim NewPT() As Byte
        NewPT = t.TransformFinalBlock(CT, 0, UBound(CT) + 1)
        
        Assert.AreEqual PT, NewPT, "Entry: " & c
    Loop
    
    WriteLine "Rijndael CFB TransformFinalBlock Decryption: " & c & " combinations tested."
    Exit Sub
    
errTrap:
    Dim en As Long
    Dim ed As String
    Dim es As String
    en = Err.Number
    ed = Err.Description
    es = Err.Source
    Debug.Print ed
    Debug.Print "Entry: " & c
    Debug.Print "PT = " & ConvertBytesToHex(PT)
    Debug.Print "CT = " & ConvertBytesToHex(CT)
    Debug.Print "NewPT = " & ConvertBytesToHex(NewPT)
    Err.Raise en, es, ed
End Sub

Public Sub TestTransformBlock()
    Dim c As Long
    On Error GoTo errTrap
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
            Assert.AreEqual (UBound(CT) + 1) - (csp.FeedbackSize \ 8), AmountTransformed, "Wrong amount transformed."
            If AmountTransformed = 0 Then
                NewPT = cArray.CreateInstance(ciByte)
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
    Exit Sub
    
errTrap:
    Dim en As Long
    Dim ed As String
    Dim es As String
    en = Err.Number
    ed = Err.Description
    es = Err.Source
    
    Debug.Print "Entry: " & c
    Debug.Print "PT = " & ConvertBytesToHex(PT)
    Debug.Print "CT = " & ConvertBytesToHex(CT)
    Debug.Print "NewPT = " & ConvertBytesToHex(NewPT)
    Err.Raise en, es, ed
End Sub

Public Sub LongTransformBlockMultipleCalls()
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
    Loop While cArray.GetLength(CT) < 32
    
    Dim NewPT() As Byte
    ReDim Preserve PT(15)
    
    Dim t As RijndaelManagedTransform
    Set t = csp.CreateDecryptor
    
    Dim i As Long
    For i = 1 To 5
        ReDim NewPT(31)
        Assert.AreEqual 16, t.TransformBlock(CT, 0, 32, NewPT, 0)
        ReDim Preserve NewPT(15)
        Assert.AreEqual PT, NewPT
        t.Reset
    Next i
End Sub



Private Sub PadWithZeros(ByRef PT() As Byte)
    ' We pad the plain text with zeros if necessary because the decrypted text
    ' will have the zero padding and we don't know what to remove.
    If csp.Padding = PaddingMode.Zeros Then
        If csp.Mode <> CFB Then
            If ((UBound(PT) + 1) Mod csp.BlockSize \ 8) <> 0 Then
                ReDim Preserve PT(0 To UBound(PT) + (csp.BlockSize \ 8 - ((UBound(PT) + 1) Mod csp.BlockSize \ 8)))
            End If
        Else
            If ((UBound(PT) + 1) Mod csp.FeedbackSize \ 8) <> 0 Then
                ReDim Preserve PT(0 To UBound(PT) + (csp.FeedbackSize \ 8 - ((UBound(PT) + 1) Mod csp.FeedbackSize \ 8)))
            End If
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
            Case "md": csp.Mode = CipherMode.CFB ' CInt(Parts(1))
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

Private Function ConvertBytesToHex(ByRef Bytes() As Byte) As String
    Dim i As Long
    Dim sb As New StringBuilder
    
    For i = 0 To UBound(Bytes)
        sb.AppendFormat "{0:X2}", Bytes(i)
    Next i
    
    ConvertBytesToHex = sb.ToString
End Function



