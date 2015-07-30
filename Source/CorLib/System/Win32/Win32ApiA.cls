VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiA"
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
' Module: Win32ApiA
'
Option Explicit
Implements IWin32Api


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindDataAToFindData(ByRef Source As WIN32_FIND_DATAA, ByRef Dest As WIN32_FIND_DATA)
    With Dest
        .cAlternateFileName = SysAllocString(StrPtr(StrConv(Source.cAlternateFileName, vbUnicode)))
        .cFileName = SysAllocString(StrPtr(StrConv(Source.cFileName, vbUnicode)))
        .dwFileAttributes = Source.dwFileAttributes
        .ftCreationTime = Source.ftCreationTime
        .ftLastAccessTime = Source.ftLastAccessTime
        .ftLastWriteTime = Source.ftLastWriteTime
        .nFileSizeHigh = Source.nFileSizeHigh
        .nFileSizeLow = Source.nFileSizeLow
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IWin32API Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IWin32API_CopyFile(ExistingFileName As String, NewFileName As String, ByVal FailIfExists As Boolean) As Long
    IWin32API_CopyFile = CopyFileA(ExistingFileName, NewFileName, FailIfExists)
End Function

Private Function IWin32API_CreateDirectory(PathName As String, Optional ByVal lpSecurityAttributes As Long = 0&) As Long
    IWin32API_CreateDirectory = CreateDirectoryA(PathName, ByVal lpSecurityAttributes)
End Function

Private Function IWin32API_CreateFile(FileName As String, ByVal DesiredAccess As Long, ByVal ShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal CreationDisposition As Long, ByVal FlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    IWin32API_CreateFile = CreateFileA(FileName, DesiredAccess, ShareMode, ByVal lpSecurityAttributes, CreationDisposition, FlagsAndAttributes, hTemplateFile)
End Function

Private Function IWin32API_CreateFileMapping(ByVal hFile As Long, ByVal lpSecurityAttributes As Long, ByVal flProtect As Long, ByVal MaximumSizeHigh As Long, ByVal MaximumSizeLow As Long, Name As String) As Long
    IWin32API_CreateFileMapping = CreateFileMappingA(hFile, ByVal lpSecurityAttributes, flProtect, MaximumSizeHigh, MaximumSizeLow, Name)
End Function

Private Function IWin32API_DeleteFile(FileName As String) As Long
    IWin32API_DeleteFile = DeleteFileA(FileName)
End Function

Private Function IWin32API_ExpandEnvironmentStrings(lpSrc As String, lpDst As String, ByVal nSize As Long) As Long
    IWin32API_ExpandEnvironmentStrings = ExpandEnvironmentStringsA(lpSrc, lpDst, nSize)
End Function

Private Function IWin32API_FindFirstFile(FileName As String, FindFileData As CorType.WIN32_FIND_DATA) As Long
    Dim Data As WIN32_FIND_DATAA
    IWin32API_FindFirstFile = FindFirstFileA(FileName, Data)
    FindDataAToFindData Data, FindFileData
End Function

Private Function IWin32API_FindNextFile(ByVal Handle As Long, FindFileData As CorType.WIN32_FIND_DATA) As Long
    Dim Data As WIN32_FIND_DATAA
    IWin32API_FindNextFile = FindNextFileA(Handle, Data)
    FindDataAToFindData Data, FindFileData
End Function

Private Function IWin32API_GetComputerName(Buffer As String, Size As Long) As Long
    IWin32API_GetComputerName = GetComputerNameA(Buffer, Size)
End Function

Private Function IWin32API_GetCurrentDirectory(ByVal BufferLength As Long, Buffer As String) As Long
    IWin32API_GetCurrentDirectory = GetCurrentDirectoryA(BufferLength, Buffer)
End Function

Private Function IWin32API_GetEnvironmentStrings() As String
    Dim lpStrings As Long
    lpStrings = GetEnvironmentStringsA
    
    If lpStrings = vbNullPtr Then _
        Throw Cor.NewInvalidOperationException(GetMessage(Err.LastDllError))
    
    Dim i As Long
    i = lpStrings
    Do While MemWord(i) <> vbNullPtr
        i = i + 1
    Loop
    
    Dim b() As Byte
    ReDim b(0 To i - lpStrings - 1)
    CopyMemory b(0), ByVal lpStrings, i - lpStrings
    
    FreeEnvironmentStringsA lpStrings
    IWin32API_GetEnvironmentStrings = StrConv(b, vbUnicode)
End Function

Private Function IWin32API_GetEnvironmentVariable(Name As String, Buffer As String, ByVal Size As Long) As Long
    IWin32API_GetEnvironmentVariable = GetEnvironmentVariableA(Name, Buffer, Size)
End Function

Private Function IWin32API_GetFileAttributes(FileName As String) As Long
    IWin32API_GetFileAttributes = GetFileAttributesA(FileName)
End Function

Private Function IWin32API_GetFileAttributesEx(FileName As String, ByVal fInfoLevelId As Long, lpFileInformation As CorType.WIN32_FILE_ATTRIBUTE_DATA) As Long
    IWin32API_GetFileAttributesEx = GetFileAttributesExA(FileName, fInfoLevelId, lpFileInformation)
End Function

Private Function IWin32API_GetFullPathName(FileName As String, ByVal BufferLength As Long, Buffer As String, ByVal lpFilePart As Long) As Long
    IWin32API_GetFullPathName = GetFullPathNameA(FileName, BufferLength, Buffer, ByVal lpFilePart)
End Function

Private Function IWin32API_GetLocaleInfo(ByVal Locale As Long, ByVal LCType As Long, lpLCData As String, ByVal cchData As Long) As Long
    IWin32API_GetLocaleInfo = GetLocaleInfoA(Locale, LCType, lpLCData, cchData)
End Function

Private Function IWin32API_GetLongPathName(ShortPath As String, LongPath As String, ByVal LongPathBufferSize As Long) As Long
    IWin32API_GetLongPathName = GetLongPathNameA(ShortPath, LongPath, LongPathBufferSize)
End Function

Private Function IWin32API_GetPrivateProfileSection(AppName As String, ReturnedString As String, ByVal Size As Long, FileName As String) As Long
    IWin32API_GetPrivateProfileSection = GetPrivateProfileSectionA(AppName, ReturnedString, Size, FileName)
End Function

Private Function IWin32API_GetPrivateProfileString(AppName As String, KeyName As String, Default As String, ReturnedString As String, ByVal Size As Long, FileName As String) As Long
    IWin32API_GetPrivateProfileString = GetPrivateProfileStringA(AppName, KeyName, Default, ReturnedString, Size, FileName)
End Function

Private Function IWin32API_GetSystemDirectory(Buffer As String, ByVal Size As Long) As Long
    IWin32API_GetSystemDirectory = GetSystemDirectoryA(Buffer, Size)
End Function

Private Function IWin32API_GetTempFileName(PathName As String, PrefixString As String, ByVal Unique As Long, TempFileName As String) As Long
    IWin32API_GetTempFileName = GetTempFileNameA(PathName, PrefixString, Unique, TempFileName)
End Function

Private Function IWin32API_GetTempPath(ByVal BufferLength As Long, Buffer As String) As Long
    IWin32API_GetTempPath = GetTempPathA(BufferLength, Buffer)
End Function

Private Function IWin32API_MoveFile(ExistingFileName As String, NewFileName As String) As Long
    IWin32API_MoveFile = MoveFileA(ExistingFileName, NewFileName)
End Function

Private Function IWin32API_RegCreateKeyEx(ByVal hKey As Long, lpSubKey As String, ByVal reserved As Long, lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
    IWin32API_RegCreateKeyEx = RegCreateKeyExA(hKey, lpSubKey, reserved, lpClass, dwOptions, samDesired, lpSecurityAttributes, phkResult, lpdwDisposition)
End Function

Private Function IWin32API_RegDeleteKey(ByVal hKey As Long, lpSubKey As String) As Long
    IWin32API_RegDeleteKey = RegDeleteKeyA(hKey, lpSubKey)
End Function

Private Function IWin32API_RegDeleteValue(ByVal hKey As Long, lpValueName As String) As Long
    IWin32API_RegDeleteValue = RegDeleteValueA(hKey, lpValueName)
End Function

Private Function IWin32API_RegEnumKeyEx(ByVal hKey As Long, ByVal dwIndex As Long, lpName As String, lpcName As Long, ByVal lpReserved As Long, lpClass As String, lpcClass As Long, lpftLastWriteTime As Currency) As Long
    IWin32API_RegEnumKeyEx = RegEnumKeyExA(hKey, dwIndex, lpName, lpcName, lpReserved, lpClass, lpcClass, lpftLastWriteTime)
End Function

Private Function IWin32API_RegEnumValue(ByVal hKey As Long, ByVal dwIndex As Long, lpValueName As String, lpcValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long
    IWin32API_RegEnumValue = RegEnumValueA(hKey, dwIndex, lpValueName, lpcValueName, lpReserved, lpType, lpData, lpcbData)
End Function

Private Function IWin32API_RegOpenKeyEx(ByVal hKey As Long, SubKey As String, ByVal Options As Long, ByVal Desired As Long, Result As Long) As Long
    IWin32API_RegOpenKeyEx = RegOpenKeyExA(hKey, SubKey, Options, Desired, Result)
End Function

Private Function IWin32API_RegQueryInfoKey(ByVal hKey As Long, lpClass As String, ByVal lpcClass As Long, ByVal lpReserve As Long, ByVal lpcSubKeys As Long, ByVal lpcMaxSubKeyLen As Long, ByVal lpcMaxClassLen As Long, ByVal lpcValues As Long, ByVal lpMaxValueNameLen As Long, ByVal lpMaxValueLen As Long, ByVal lpSecurityDescriptor As Long, ByVal lpLastWriteTime As Long) As Long
    IWin32API_RegQueryInfoKey = RegQueryInfoKeyA(hKey, lpClass, ByVal lpcClass, lpReserve, ByVal lpcSubKeys, ByVal lpcMaxSubKeyLen, ByVal lpcMaxClassLen, ByVal lpcValues, ByVal lpMaxValueNameLen, ByVal lpMaxValueLen, ByVal lpSecurityDescriptor, ByVal lpLastWriteTime)
End Function

Private Function IWin32API_RegQueryValueEx(ByVal hKey As Long, ValueName As String, ValueType As Long, ByVal lpData As Long, lpcbData As Long) As Long
    IWin32API_RegQueryValueEx = RegQueryValueExA(hKey, ValueName, 0, ValueType, ByVal lpData, lpcbData)
End Function

Private Function IWin32API_RegQueryValueExStr(ByVal hKey As Long, ValueName As String, ValueType As Long, Data As String) As Long
    Dim Bytes() As Byte
    ReDim Bytes(0 To Len(Data) - 1)
    
    Dim Result As Long
    Result = RegQueryValueExA(hKey, ValueName, 0, ValueType, ByVal VarPtr(Bytes(0)), Len(Data))
    
    If Result = ERROR_SUCCESS Then
        Data = StrConv(Bytes, vbUnicode)
    End If
    IWin32API_RegQueryValueExStr = Result
End Function

Private Function IWin32API_RegSetValueEx(ByVal hKey As Long, ValueName As String, ByVal ValueType As Long, ByVal lpData As Long, ByVal cbData As Long) As Long
    IWin32API_RegSetValueEx = RegSetValueExA(hKey, ValueName, 0, ValueType, lpData, cbData)
End Function

Private Function IWin32API_RegSetValueExStr(ByVal hKey As Long, ValueName As String, ByVal ValueType As Long, Data As String) As Long
    Dim Bytes() As Byte
    Bytes = StrConv(Data, vbFromUnicode)
    IWin32API_RegSetValueExStr = RegSetValueExA(hKey, ValueName, 0, ValueType, VarPtr(Bytes(0)), Len(Data))
End Function

Private Function IWin32API_RemoveDirectory(lpPathName As String) As Long
    IWin32API_RemoveDirectory = RemoveDirectoryA(lpPathName)
End Function

Private Function IWin32API_SetCurrentDirectory(PathName As String) As Long
    IWin32API_SetCurrentDirectory = SetCurrentDirectoryA(PathName)
End Function

Private Function IWin32API_SetEnvironmentVariable(Name As String, Value As String) As Long
    IWin32API_SetEnvironmentVariable = SetEnvironmentVariableA(Name, Value)
End Function

Private Function IWin32API_SetFileAttributes(FileName As String, ByVal dwFileAttributes As Long) As Long
    IWin32API_SetFileAttributes = SetFileAttributesA(FileName, dwFileAttributes)
End Function

Private Function IWin32API_WritePrivateProfileSection(AppName As String, Value As String, FileName As String) As Long
    IWin32API_WritePrivateProfileSection = WritePrivateProfileSectionA(AppName, Value, FileName)
End Function

Private Function IWin32API_WritePrivateProfileString(AppName As String, KeyName As String, Value As String, FileName As String) As Long
    IWin32API_WritePrivateProfileString = WritePrivateProfileStringA(AppName, KeyName, Value, FileName)
End Function
