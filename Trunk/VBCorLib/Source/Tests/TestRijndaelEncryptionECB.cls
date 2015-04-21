VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRijndaelEncryptionECB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRijndaelEncryptionECB
'
Option Explicit

Private csp As RijndaelManaged
Private FF As Long
Private PT() As Byte
Private CT() As Byte

Public Sub Setup()
    Set csp = New RijndaelManaged
    FF = FreeFile
    Open App.Path & "\Tests\Cryptography Files\Rijndael_ECB_Entries.txt" For Input As #FF
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
        c = c + 1
        
        Dim NewCT() As Byte
        Dim t As ICryptoTransform
        Set t = csp.CreateEncryptor
        NewCT = t.TransformFinalBlock(PT, 0, cArray.GetLength(PT))
                
        ' We need atleast 2 bytes of padding for ISO10126. 1 for the number of padded bytes,
        ' and one as the random value filling the rest of the padded bytes.
        If (csp.Padding = PaddingMode.ISO10126) And (cArray.GetLength(CT) - cArray.GetLength(PT) > 1) Then
            ' For padding mode ISO10126, the best we can do without decrypting
            ' the text is to make sure the same encrypted text has a different output.
            ' However, there is the possibility the same random fill bytes will be generated.
            ' So this test may fail. It should be run again to ensure the failure is genuine.
            
            Assert.That NewCT, Iz.Not.EqualTo(t.TransformFinalBlock(PT, 0, cArray.GetLength(PT))), "Entry " & c & ". Try running test again."
        Else
            Assert.AreEqual CT, NewCT, "Entry " & c & " failed. PT = " & ConvertBytesToHex(PT)
        End If
    Loop
    Exit Sub
    
errTrap:
    Debug.Print "PT = " & ConvertBytesToHex(PT)
    Debug.Print "CT = " & ConvertBytesToHex(CT)
    Debug.Print "NewCT = " & ConvertBytesToHex(NewCT)
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub TestTransformBlock()
    Dim c As Long
    
    Do While GetNextEntry
        c = c + 1
        
        ' We skip this kind of padding mode because it is random
        ' so we can't reproduce the same result as in the file.
        ' Special tests will be needed for this kind of padding.
        If csp.Padding <> PaddingMode.ISO10126 Then
            Dim AmountTransformed As Long
            Dim NewCT() As Byte
            Dim t As ICryptoTransform
            Set t = csp.CreateEncryptor
            
            ReDim NewCT(UBound(CT))
            
            Dim FullBlocks As Long
            FullBlocks = cArray.GetLength(PT) \ t.InputBlockSize
            If FullBlocks > 0 Then
                AmountTransformed = t.TransformBlock(PT, 0, FullBlocks * t.InputBlockSize, NewCT, 0)
                Assert.AreEqual FullBlocks * t.InputBlockSize, AmountTransformed, "Entry " & c & " is wrong amount transformed."
            End If
            
            Dim FinalBlock() As Byte
            FinalBlock = t.TransformFinalBlock(PT, FullBlocks * t.InputBlockSize, cArray.GetLength(PT) - FullBlocks * t.InputBlockSize)
            
            Buffer.BlockCopy FinalBlock, 0, NewCT, FullBlocks * t.OutputBlockSize, cArray.GetLength(FinalBlock)
            
            Assert.AreEqual CT, NewCT, "Entry " & c & " is wrong text."
        End If
    Loop
End Sub

Public Sub ResetForNewText()
    Do
        GetNextEntry
    Loop While (cArray.GetLength(PT) < 24) Or (csp.Padding = ISO10126)
    
    Dim NewCT() As Byte
    ReDim NewCT(15)
    ReDim Preserve CT(15)
    
    Dim t As RijndaelManagedTransform
    Set t = csp.CreateEncryptor
    
    Dim i As Long
    For i = 1 To 5
        ReDim NewCT(15)
        Assert.AreEqual 16, t.TransformBlock(PT, 0, 16, NewCT, 0)
        Assert.AreEqual CT, NewCT
        t.Reset
    Next i
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

Private Function ConvertBytesToHex(ByRef Bytes() As Byte) As String
    Dim i As Long
    Dim sb As New StringBuilder
    
    For i = 0 To UBound(Bytes)
        sb.AppendFormat "{0:X2}", Bytes(i)
    Next i
    
    ConvertBytesToHex = sb.ToString
End Function

