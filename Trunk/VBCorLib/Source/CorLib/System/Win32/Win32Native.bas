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

Private Type FileNameBuffer
    Buffer As String * 32000
End Type

Public Api As IWin32Api
Private FileName As FileNameBuffer

Public Sub InitWin32Api()
    Set Api = New Win32ApiW
End Sub

Public Function SafeCreateFile(FileName As String, ByVal DesiredAccess As FileAccess, ByVal ShareMode As FileShare, ByVal CreationDisposition As FileMode, Optional ByVal FlagsAndAttributes = FILE_ATTRIBUTE_NORMAL) As SafeFileHandle
    Dim FileHandle As Long
    FileHandle = CreateFileW(FileName, DesiredAccess, ShareMode, ByVal 0, CreationDisposition, FlagsAndAttributes, 0)
    Set SafeCreateFile = Cor.NewSafeFileHandle(FileHandle, True)
End Function

Public Function SafeFindFirstFile(ByRef FileName As String, ByRef FindFileData As WIN32_FIND_DATA) As SafeFindHandle
    Dim WideData    As WIN32_FIND_DATAW
    Dim FileHandle  As Long
    
    FileHandle = FindFirstFileW(FileName, WideData)
    FindDataWToFindData WideData, FindFileData
    Set SafeFindFirstFile = Cor.NewSafeFindHandle(FileHandle, True)
End Function

Public Function GetModuleFileName(ByVal hModule As Long, ByRef lpFileName As String, ByRef nSize As Long) As Long
    GetModuleFileName = GetModuleFileNameW(hModule, lpFileName, nSize)
End Function

Public Function GetUserNameEx(ByVal NameFormat As Long, ByRef lpNameBuffer As String, ByRef nSize As Long) As Long
    GetUserNameEx = GetUserNameExW(NameFormat, lpNameBuffer, nSize)
End Function

Public Function LookupAccountName(ByVal lpSystemName As String, ByVal lpAccountName As String, ByRef Sid As String, ByRef cbSid As Long, ByRef ReferencedDomainName As String, ByRef cbReferencedDomainName As Long, ByRef peUse As Long) As Long
    LookupAccountName = LookupAccountNameW(lpSystemName, lpAccountName, StrPtr(Sid), cbSid, ReferencedDomainName, cbReferencedDomainName, peUse)
End Function

Public Function GetProcessWindowStation() As Long
    GetProcessWindowStation = VBCorType.GetProcessWindowStation
End Function

Public Function GetUserObjectInformation(ByVal hObj As Long, ByVal nIndex As Long, ByVal pvInfo As Long, ByVal nLength As Long, ByRef lpnLengthNeeded As Long) As Long
    GetUserObjectInformation = GetUserObjectInformationW(hObj, nIndex, pvInfo, nLength, lpnLengthNeeded)
End Function

Public Function GetSystemMenu(ByVal hwnd As Long, ByVal bRevert As Boolean) As Long
    Dim BoolRevert As BOOL
    BoolRevert = IIf(bRevert, BOOL_TRUE, BOOL_FALSE)
    GetSystemMenu = VBCorType.GetSystemMenu(hwnd, BoolRevert)
End Function

Public Function RemoveMenu(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    RemoveMenu = VBCorType.RemoveMenu(hMenu, nPosition, wFlags)
End Function

Public Function SetCurrentDirectory(ByRef PathName As String) As Boolean
    SetCurrentDirectory = (SetCurrentDirectoryW(PathName) <> BOOL_FALSE)
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
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








