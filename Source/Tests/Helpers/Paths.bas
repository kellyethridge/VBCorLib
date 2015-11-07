Attribute VB_Name = "Paths"
Option Explicit

Public Property Get CryptographyFolder() As String
    CryptographyFolder = App.Path & "\Tests\Cryptography Files"
End Property

Public Function GetCryptoPath(ByRef FileName As String) As String
    GetCryptoPath = CryptographyFolder & "\" & FileName
End Function

Public Function Missing(Optional ByRef Value As Variant) As Variant
    Missing = Value
End Function
