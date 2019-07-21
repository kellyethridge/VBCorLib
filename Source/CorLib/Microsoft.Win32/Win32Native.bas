Attribute VB_Name = "Win32Native"
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
' Module: Win32Native
'

' These are here because these are not supported on Win9x.
Option Explicit

Public Const CSIDL_ADMINTOOLS               As Long = &H30
Public Const CSIDL_CDBURN_AREA              As Long = &H3B
Public Const CSIDL_COMMON_ADMINTOOLS        As Long = &H2F
Public Const CSIDL_COMMON_DOCUMENTS         As Long = &H2E
Public Const CSIDL_COMMON_MUSIC             As Long = &H35
Public Const CSIDL_COMMON_OEM_LINKS         As Long = &H3A
Public Const CSIDL_COMMON_PICTURES          As Long = &H36
Public Const CSIDL_COMMON_STARTMENU         As Long = &H16
Public Const CSIDL_COMMON_PROGRAMS          As Long = &H17
Public Const CSIDL_COMMON_STARTUP           As Long = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY  As Long = &H19
Public Const CSIDL_COMMON_TEMPLATES         As Long = &H2D
Public Const CSIDL_COMMON_VIDEO             As Long = &H37
Public Const CSIDL_FONTS                    As Long = &H14
Public Const CSIDL_MYVIDEO                  As Long = &HE
Public Const CSIDL_NETHOOD                  As Long = &H13
Public Const CSIDL_PRINTHOOD                As Long = &H1B
Public Const CSIDL_PROFILE                  As Long = &H28
Public Const CSIDL_PROGRAM_FILES_COMMONX86  As Long = &H2C
Public Const CSIDL_PROGRAM_FILESX86         As Long = &H2A
Public Const CSIDL_RESOURCES                As Long = &H38
Public Const CSIDL_RESOURCES_LOCALIZED      As Long = &H39
Public Const CSIDL_SYSTEMX86                As Long = &H29
Public Const CSIDL_WINDOWS                  As Long = &H24
Public Const CSIDL_APPDATA                  As Long = &H1A
Public Const CSIDL_COMMON_APPDATA           As Long = &H23
Public Const CSIDL_LOCAL_APPDATA            As Long = &H1C
Public Const CSIDL_COOKIES                  As Long = &H21
Public Const CSIDL_FAVORITES                As Long = &H6
Public Const CSIDL_HISTORY                  As Long = &H22
Public Const CSIDL_INTERNET_CACHE           As Long = &H20
Public Const CSIDL_PROGRAMS                 As Long = &H2
Public Const CSIDL_RECENT                   As Long = &H8
Public Const CSIDL_SENDTO                   As Long = &H9
Public Const CSIDL_STARTMENU                As Long = &HB
Public Const CSIDL_STARTUP                  As Long = &H7
Public Const CSIDL_SYSTEM                   As Long = &H25
Public Const CSIDL_TEMPLATES                As Long = &H15
Public Const CSIDL_DESKTOPDIRECTORY         As Long = &H10
Public Const CSIDL_PERSONAL                 As Long = &H5
Public Const CSIDL_PROGRAM_FILES            As Long = &H26
Public Const CSIDL_PROGRAM_FILES_COMMON     As Long = &H2B
Public Const CSIDL_DESKTOP                  As Long = &H0
Public Const CSIDL_DRIVES                   As Long = &H11
Public Const CSIDL_MYMUSIC                  As Long = &HD
Public Const CSIDL_MYPICTURES               As Long = &H27

Private Type FileNameBuffer
    Buffer As String * 32000
End Type

Private FileName As FileNameBuffer

Public Function MakeHRFromErrorCode(ByVal ErrorCode As Long)
    MakeHRFromErrorCode = &H80070000 Or ErrorCode
End Function

Public Function SafeCreateFile(FileName As String, ByVal DesiredAccess As FileAccess, ByVal ShareMode As FileShare, ByVal CreationDisposition As FileMode, Optional ByVal FlagsAndAttributes = FILE_ATTRIBUTE_NORMAL) As SafeFileHandle
    Dim AccessFlag As Long
    Select Case DesiredAccess
        Case FileAccess.ReadAccess
            AccessFlag = GENERIC_READ
        Case FileAccess.WriteAccess
            AccessFlag = GENERIC_WRITE
        Case FileAccess.ReadWriteAccess
            AccessFlag = GENERIC_READ Or GENERIC_WRITE
        Case Else
            Error.ArgumentOutOfRange "DesiredAccess", ArgumentOutOfRange_Enum
    End Select
    
    Dim FileHandle As Long
    Dim SafeHandle As SafeFileHandle
    FileHandle = CreateFileW(FileName, AccessFlag, ShareMode, ByVal 0, CreationDisposition, FlagsAndAttributes, 0)
    Set SafeHandle = Cor.NewSafeFileHandle(FileHandle, True)
    
    If Not SafeHandle.IsInvalid Then
        Dim FileType As Long
        FileType = GetFileType(SafeHandle)
        
        If FileType <> FILE_TYPE_DISK Then _
            Error.NotSupported NotSupported_FileStreamOnNonFiles
    End If
    
    Set SafeCreateFile = SafeHandle
End Function

Public Function SafeFindFirstFile(ByRef FileName As String, ByRef FindFileData As WIN32_FIND_DATAW) As SafeFindHandle
    Dim FileHandle  As Long
    
    FileHandle = FindFirstFileW(FileName, FindFileData)
    Set SafeFindFirstFile = Cor.NewSafeFindHandle(FileHandle, True)
End Function

Public Function SafeCreateFileMapping(ByVal FileHandle As Long, ByVal Protect As Long, ByVal CapacityHigh As Long, ByVal CapacityLow As Long, ByRef MapName As String) As SafeFileHandle
    Dim Handle As Long
    
    Handle = CreateFileMappingW(FileHandle, ByVal 0&, Protect, CapacityHigh, CapacityLow, MapName)
    Set SafeCreateFileMapping = Cor.NewSafeFileHandle(Handle, True)
End Function

Public Function GetSystemMenu(ByVal hwnd As Long, ByVal bRevert As Boolean) As Long
    Dim BoolRevert As BOOL
    BoolRevert = IIfLong(bRevert, BOOL_TRUE, BOOL_FALSE)
    GetSystemMenu = CorType.GetSystemMenu(hwnd, BoolRevert)
End Function

Public Function GetMessage(ByVal ErrorCode As Long) As String
    Const FORMAT_MESSAGE_FLAGS As Long = FORMAT_MESSAGE_FROM_SYSTEM Or _
                                         FORMAT_MESSAGE_IGNORE_INSERTS Or _
                                         FORMAT_MESSAGE_ARGUMENT_ARRAY
    Dim Buf     As String
    Dim Size    As Long
    
    Buf = String$(1024, vbNullChar)
    Size = FormatMessageA(FORMAT_MESSAGE_FLAGS, ByVal 0&, ErrorCode, 0, Buf, Len(Buf), ByVal 0&)
    
    If Size > 0 Then
        GetMessage = Left$(Buf, Size - 2)
    Else
        GetMessage = Environment.GetResourceString(UnknownError_Num, ErrorCode)
    End If
End Function







