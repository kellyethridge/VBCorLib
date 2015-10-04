Attribute VB_Name = "ConversionHelper"
Option Explicit

Public Function HexToBytes(ByRef s As String) As Byte()
    s = Replace$(s, " ", "")
    Dim Bytes() As Byte
    ReDim Bytes(0 To Len(s) \ 2 - 1)
    
    Dim i As Long
    For i = 0 To UBound(Bytes)
        Bytes(i) = CByte("&h" & Mid$(s, (i * 2) + 1, 2))
    Next i
    
    HexToBytes = Bytes
End Function
