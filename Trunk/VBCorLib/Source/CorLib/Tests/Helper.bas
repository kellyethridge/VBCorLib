Attribute VB_Name = "Helper"
Option Explicit

Public Function FolderExists(ByRef Folder As String) As Boolean
    FolderExists = Len(Dir$(Folder, vbDirectory)) > 0
End Function

Public Sub DeleteFolder(ByRef Folder As String)
    If FolderExists(Folder) Then
        RmDir Folder
    End If
End Sub

Public Sub CreateFolder(ByRef Folder As String)
    If Not FolderExists(Folder) Then
        MkDir Folder
    End If
End Sub

Public Function FileExists(ByRef FileName As String) As Boolean
    FileExists = Len(Dir$(FileName, vbNormal)) > 0
End Function

Public Sub CreateFile(ByRef FileName As String)
    If Not FileExists(FileName) Then
        Dim FileNumber As Long
        FileNumber = FreeFile
        Open FileName For Output As #FileNumber
        Close FileNumber
    End If
End Sub

Public Sub DeleteFile(ByRef FileName As String)
    If FileExists(FileName) Then
        Kill FileName
    End If
End Sub

