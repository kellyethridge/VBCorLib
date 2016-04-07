VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: Win32ApiW
'
Option Explicit
Implements IWin32Api


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindDataWToFindData(ByRef Source As WIN32_FIND_DATAW, ByRef Dest As WIN32_FIND_DATA)
    With Dest
        .cAlternateFileName = SysAllocString(VarPtr(Source.cAlternateFileName(0)))
        .cFileName = SysAllocString(VarPtr(Source.cFileName(0)))
        .dwFileAttributes = Source.dwFileAttributes
        .ftCreationTime = Source.ftCreationTime
        .ftLastAccessTime = Source.ftLastAccessTime
        .ftLastWriteTime = Source.ftLastWriteTime
        .nFileSizeHigh = Source.nFileSizeHigh
        .nFileSizeLow = Source.nFileSizeLow
    End With
End Sub

Private Function MakeWide(ByVal s As String) As String
    s = Path.GetFullPath(s)
    
    ' Check if we have a UNC path.
    Dim root As String
    If Left$(Path.GetPathRoot(s), 2) = "\\" Then
        ' MSDN says the format is \\?\UNC\Server\Share\... ,
        ' so we need to trim off the first backslash from the path
        s = Mid$(s, 2)
        root = "UNC"
    End If
    
    MakeWide = "\\?\" & root & s
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IWin32API Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IWin32API_CopyFile(ExistingFileName As String, NewFileName As String, ByVal FailIfExists As Boolean) As Long
    IWin32API_CopyFile = CopyFileW(MakeWide(ExistingFileName), MakeWide(NewFileName), FailIfExists)
End Function

Private Function IWin32API_CreateDirectory(PathName As String, Optional ByVal lpSecurityAttributes As Long = 0&) As Long
    IWin32API_CreateDirectory = CreateDirectoryW(MakeWide(PathName), ByVal lpSecurityAttributes)
End Function

Private Function IWin32API_CreateFile(FileName As String, ByVal DesiredAccess As Long, ByVal ShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal CreationDisposition As Long, ByVal FlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    IWin32API_CreateFile = CreateFileW(MakeWide(FileName), DesiredAccess, ShareMode, ByVal lpSecurityAttributes, CreationDisposition, FlagsAndAttributes, hTemplateFile)
End Function

Private Function IWin32API_CreateFileMapping(ByVal hFile As Long, ByVal lpSecurityAttributes As Long, ByVal flProtect As Long, ByVal MaximumSizeHigh As Long, ByVal MaximumSizeLow As Long, Name As String) As Long
    IWin32API_CreateFileMapping = CreateFileMappingW(hFile, ByVal lpSecurityAttributes, flProtect, MaximumSizeHigh, MaximumSizeLow, Name)
End Function

Private Function IWin32API_DeleteFile(FileName As String) As Long
    IWin32API_DeleteFile = DeleteFileW(MakeWide(FileName))
End Function

Private Function IWin32API_ExpandEnvironmentStrings(lpSrc As String, lpDst As String, ByVal nSize As Long) As Long
    IWin32API_ExpandEnvironmentStrings = ExpandEnvironmentStringsW(lpSrc, lpDst, nSize)
End Function

Private Function IWin32API_FindFirstFile(FileName As String, FindFileData As CorType.WIN32_FIND_DATA) As Long
    Dim Data As WIN32_FIND_DATAW
    IWin32API_FindFirstFile = FindFirstFileW(MakeWide(FileName), Data)
    FindDataWToFindData Data, FindFileData
End Function

Private Function IWin32API_FindNextFile(ByVal Handle As Long, FindFileData As CorType.WIN32_FIND_DATA) As Long
    Dim Data As WIN32_FIND_DATAW
    IWin32API_FindNextFile = FindNextFileW(Handle, Data)
    FindDataWToFindData Data, FindFileData
End Function

Private Function IWin32API_GetComputerName(Buffer As String, Size As Long) As Long
    IWin32API_GetComputerName = GetComputerNameW(Buffer, Size)
End Function

Private Function IWin32API_GetCurrentDirectory(ByVal BufferLength As Long, Buffer As String) As Long
    IWin32API_GetCurrentDirectory = GetCurrentDirectoryW(BufferLength, Buffer)
End Function

Private Function IWin32API_GetEnvironmentStrings() As String
    Dim lpStrings As Long
    lpStrings = GetEnvironmentStringsW
    
    If lpStrings = vbNullPtr Then _
        Throw Cor.NewInvalidOperationException(GetMessage(Err.LastDllError))
        
    Dim i As Long
    i = lpStrings
    Do While MemLong(i) <> vbNullPtr
        i = i + 1
    Loop
    
    Dim s As String
    s = SysAllocStringLen(lpStrings, (i - lpStrings) \ 2)
    
    FreeEnvironmentStringsW lpStrings
    IWin32API_GetEnvironmentStrings = s
End Function

Private Function IWin32API_GetEnvironmentVariable(Name As String, Buffer As String, ByVal Size As Long) As Long
    IWin32API_GetEnvironmentVariable = GetEnvironmentVariableW(Name, Buffer, Size)
End Function

Private Function IWin32API_GetFileAttributes(FileName As String) As Long
    IWin32API_GetFileAttributes = GetFileAttributesW(FileName)
End Function

Private Function IWin32API_GetFileAttributesEx(FileName As String, ByVal fInfoLevelId As Long, lpFileInformation As CorType.WIN32_FILE_ATTRIBUTE_DATA) As Long
    IWin32API_GetFileAttributesEx = GetFileAttributesExW(FileName, fInfoLevelId, lpFileInformation)
End Function

Private Function IWin32API_GetFullPathName(FileName As String, ByVal BufferLength As Long, Buffer As String, ByVal lpFilePart As Long) As Long
    IWin32API_GetFullPathName = GetFullPathNameW(FileName, BufferLength, Buffer, ByVal lpFilePart)
End Function

Private Function IWin32API_GetLocaleInfo(ByVal Locale As Long, ByVal LCType As Long, lpLCData As String, ByVal cchData As Long) As Long
    IWin32API_GetLocaleInfo = GetLocaleInfoW(Locale, LCType, lpLCData, cchData)
End Function

Private Function IWin32API_GetLongPathName(ShortPath As String, LongPath As String, ByVal LongPathBufferSize As Long) As Long
    IWin32API_GetLongPathName = GetLongPathNameW(ShortPath, LongPath, LongPathBufferSize)
End Function

Private Function IWin32API_GetPrivateProfileSection(AppName As String, ReturnedString As String, ByVal Size As Long, FileName As String) As Long
    IWin32API_GetPrivateProfileSection = GetPrivateProfileSectionW(AppName, ReturnedString, Size, MakeWide(FileName))
End Function

Private Function IWin32API_GetPrivateProfileString(AppName As String, KeyName As String, Default As String, ReturnedString As String, ByVal Size As Long, FileName As String) As Long
    IWin32API_GetPrivateProfileString = GetPrivateProfileStringW(AppName, KeyName, Default, ReturnedString, Size, MakeWide(FileName))
End Function

Private Function IWin32API_GetSystemDirectory(Buffer As String, ByVal Size As Long) As Long
    IWin32API_GetSystemDirectory = GetSystemDirectoryW(Buffer, Size)
End Function

Private Function IWin32API_GetTempFileName(PathName As String, PrefixString As String, ByVal Unique As Long, TempFileName As String) As Long
    IWin32API_GetTempFileName = GetTempFileNameW(PathName, PrefixString, Unique, TempFileName)
End Function

Private Function IWin32API_GetTempPath(ByVal BufferLength As Long, Buffer As String) As Long
    IWin32API_GetTempPath = GetTempPathW(BufferLength, Buffer)
End Function

Private Function IWin32API_MoveFile(ExistingFileName As String, NewFileName As String) As Long
    IWin32API_MoveFile = MoveFileW(MakeWide(ExistingFileName), MakeWide(NewFileName))
End Function

Private Function IWin32API_RemoveDirectory(lpPathName As String) As Long
    IWin32API_RemoveDirectory = RemoveDirectoryW(MakeWide(lpPathName))
End Function

Private Function IWin32API_SetCurrentDirectory(PathName As String) As Long
    IWin32API_SetCurrentDirectory = SetCurrentDirectoryW(PathName)
End Function

Private Function IWin32API_SetEnvironmentVariable(Name As String, Value As String) As Long
    IWin32API_SetEnvironmentVariable = SetEnvironmentVariableW(Name, Value)
End Function

Private Function IWin32API_SetFileAttributes(FileName As String, ByVal dwFileAttributes As Long) As Long
    IWin32API_SetFileAttributes = SetFileAttributesW(FileName, dwFileAttributes)
End Function

Private Function IWin32API_WritePrivateProfileSection(AppName As String, Value As String, FileName As String) As Long
    IWin32API_WritePrivateProfileSection = WritePrivateProfileSectionW(AppName, Value, MakeWide(FileName))
End Function

Private Function IWin32API_WritePrivateProfileString(AppName As String, KeyName As String, Value As String, FileName As String) As Long
    IWin32API_WritePrivateProfileString = WritePrivateProfileStringW(AppName, KeyName, Value, MakeWide(FileName))
End Function

