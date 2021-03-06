Attribute VB_Name = "ConversionHelper"
Option Explicit

Public Function HexToBytes(ByRef s As String, Optional ByVal Reverse As Boolean = False) As Byte()
    s = Replace$(s, " ", "")
    
    If Len(s) = 0 Then
        HexToBytes = Cor.NewBytes()
        Exit Function
    End If
    
    Dim Bytes() As Byte
    ReDim Bytes(0 To Len(s) \ 2 - 1)
    
    Dim i As Long
    For i = 0 To UBound(Bytes)
        Bytes(i) = CByte("&h" & Mid$(s, (i * 2) + 1, 2))
    Next i
    
    If Reverse Then
        CorArray.Reverse Bytes
    End If
    
    HexToBytes = Bytes
End Function

Public Function BytesToHex(ByRef Bytes() As Byte) As String
    Dim i As Long
    
    For i = LBound(Bytes) To UBound(Bytes)
        BytesToHex = BytesToHex & Right$("0" & Hex$(Bytes(i)), 2)
    Next
End Function

Public Function RepeatString(ByVal Pattern As String, ByVal Count As Long) As String
    Dim sb As New SimplyVBComp.StringBuilder
    Dim i As Long
    For i = 1 To Count
        sb.Append Pattern
    Next
    RepeatString = sb.ToString
End Function

Public Function HexString(ByVal Value As Byte, ByVal Count As Long) As String
    HexString = RepeatString(Right$(Hex$(Value), 2), Count)
End Function

Public Function TextToHex(ByRef s As String) As String
    Dim Bytes() As Byte
    Bytes = Encoding.ASCII.GetBytes(s)
    
    Dim sb As New SimplyVBComp.StringBuilder
    Dim i As Long
    For i = 0 To UBound(Bytes)
        sb.Append Right$(Hex$(Bytes(i)), 2)
    Next
    
    TextToHex = sb.ToString
End Function
