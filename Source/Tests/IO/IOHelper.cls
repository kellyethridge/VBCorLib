VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCurrentDirectory   As String
Private mTempFolder         As String

Public Property Get TempFolder() As String
    TempFolder = mTempFolder
End Property

Public Property Get TempSubFolder() As String
      TempSubFolder = "VBCorlib.IO.Tests"
End Property

Public Function CreateFile(Optional ByVal FileName As String = "File.txt", Optional ByVal Data As String) As String
    CreateFile = Path.Combine(mTempFolder, FileName)
    If Not FileExists(CreateFile) Then
        Dim FileNumber As Long
        FileNumber = FreeFile
        Open CreateFile For Output As #FileNumber
        
        If Len(Data) > 0 Then
            Print #FileNumber, Data;
        End If
        
        Close #FileNumber
    End If
End Function

Public Function CreateFolder(Optional ByVal SubFolder As String = "SubFolder") As String
    CreateFolder = Path.Combine(mTempFolder, SubFolder)
    If Not FolderExists(CreateFolder) Then
        MkDir CreateFolder
    End If
End Function

Public Function ReadFile(ByVal FileName As String) As String
    Dim FileNumber As Long
    FileNumber = FreeFile
    Open FileName For Input As #FileNumber
    ReadFile = Input(LOF(FileNumber), FileNumber)
    Close #FileNumber
End Function

Public Function ReadFileBytes(ByVal FileName As String) As Byte()
    Dim FileNumber As Long
    FileNumber = FreeFile
    Open FileName For Binary As #FileNumber
    ReadFileBytes = InputB(LOF(FileNumber), FileNumber)
    Close #FileNumber
End Function

Private Sub Setup()
    mCurrentDirectory = Directory.GetCurrentDirectory
    mTempFolder = Path.Combine(Path.GetTempPath, TempSubFolder)
    
    If Not Directory.Exists(mTempFolder) Then
        Directory.CreateDirectory mTempFolder
    End If
End Sub

Private Sub Teardown()
    Directory.SetCurrentDirectory mCurrentDirectory
    
    If Directory.Exists(mTempFolder) Then
        Directory.Delete mTempFolder, True
    End If
End Sub

Private Sub Class_Initialize()
    Setup
End Sub

Private Sub Class_Terminate()
    Teardown
End Sub
