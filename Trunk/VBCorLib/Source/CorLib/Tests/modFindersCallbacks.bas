Attribute VB_Name = "modFindersCallbacks"
'
' modFindersCallbacks
'
Option Explicit

Public Function FindHelloString(ByRef s As String) As Boolean
    FindHelloString = (s = "Hello")
End Function

Public Function FindInt32(ByRef o As Int32) As Boolean
    FindInt32 = (o.mValue = 32)
End Function

Public Function FindVBGuid(ByRef g As VBGUID) As Boolean
    FindVBGuid = (g.Data1 = 7)
End Function

'
' cArray.ForEach callback
'
Public Sub SetToNumber(ByRef e As Long)
    e = 5
End Sub
