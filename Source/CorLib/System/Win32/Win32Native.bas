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

Public Api As IWin32Api

' psapi.dll
Public Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Public Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal Process As Long, ByRef ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long

Public Sub InitWin32Api()
    Set Api = New Win32ApiW
End Sub

Public Function SafeCreateFile(FileName As String, ByVal DesiredAccess As FileAccess, ByVal ShareMode As FileShare, ByVal CreationDisposition As FileMode) As SafeFileHandle
    Dim FileHandle As Long
    FileHandle = Api.CreateFile(FileName, DesiredAccess, ShareMode, ByVal 0, CreationDisposition, FILE_ATTRIBUTE_NORMAL, 0)
    Set SafeCreateFile = Cor.NewSafeFileHandle(FileHandle, True)
End Function

Public Function SafeFindFirstFile(ByRef FileName As String, ByRef FindFileData As WIN32_FIND_DATA) As SafeFindHandle
    Dim FileHandle As Long
    FileHandle = Api.FindFirstFile(FileName, FindFileData)
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

Public Function GetSystemMenu(ByVal hWnd As Long, ByVal bRevert As Boolean) As Long
    Dim boolRevert As BOOL
    boolRevert = IIf(bRevert, BOOL_TRUE, BOOL_FALSE)
    GetSystemMenu = VBCorType.GetSystemMenu(hWnd, boolRevert)
End Function

Public Function RemoveMenu(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    RemoveMenu = VBCorType.RemoveMenu(hMenu, nPosition, wFlags)
End Function












