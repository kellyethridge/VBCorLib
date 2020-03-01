Attribute VB_Name = "MiscHelper"
Option Explicit
Public Declare Function GetCalendarInfo Lib "kernel32.dll" Alias "GetCalendarInfoA" (ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long, ByVal lpCalData As String, ByVal cchData As Long, ByRef lpValue As Any) As Long

Public Const LOCALE_USER_DEFAULT As Long = 1024
Public Const CAL_GREGORIAN As Long = 1
Public Const CAL_HEBREW As Long = 8
Public Const CAL_KOREA As Long = 5
Public Const CAL_THAI As Long = 7
Public Const CAL_ITWODIGITYEARMAX As Long = &H30
Public Const LOCALE_RETURN_NUMBER As Long = &H20000000
Public Const CAL_RETURN_NUMBER As Long = LOCALE_RETURN_NUMBER

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

Public Function GenerateBytes(ByVal Size As Long) As Byte()
    Dim Ran As New Random
    Dim Result() As Byte
    ReDim Result(0 To Size - 1)
    
    Ran.NextBytes Result
    
    GenerateBytes = Result
End Function

Public Property Get NullBytes() As Byte()
End Property

Public Property Get NullChars() As Integer()
End Property

Public Property Get NullStrings() As String()
End Property
