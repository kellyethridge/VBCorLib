Attribute VB_Name = "MiscHelper"
Option Explicit

Public Function Missing(Optional ByRef Value As Variant) As Variant
    Missing = Value
End Function

Public Function NewInt32(ByVal Value As Long) As Int32
    Set NewInt32 = New Int32
    NewInt32.Init Value
End Function

Public Function GenerateString(ByVal Size As Long) As String
    Dim Ran As New Random
    Dim sb As New StringBuilder
    Dim i As Long
    
    For i = 1 To Size
        Dim Ch As Long
        Ch = Ran.NextRange(32, Asc("z"))
        sb.AppendChar Ch
    Next
    
    GenerateString = sb.ToString
End Function

Public Property Get NullBytes() As Byte()
End Property

