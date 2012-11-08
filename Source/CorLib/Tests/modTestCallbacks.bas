Attribute VB_Name = "modTestCallbacks"
'
' Place to put all function callbacks.
'
Option Explicit

Public Sub FirstLetterCopier(ByRef Target As String, ByRef Source As String)
    Target = Left$(Source, 1)
End Sub
