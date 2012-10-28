Attribute VB_Name = "modWin32Api"
'    CopyRight (c) 2006 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: modWin32Api
'

' These are here because these are not supported on Win9x.

Option Explicit

Public API As IWin32API


'
' kernel32.dll
'
Public Declare Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
Public Declare Function GetSystemDefaultUILanguage Lib "kernel32.dll" () As Long
Public Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function VirtualProtect Lib "kernel32.dll" (ByRef lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long

' secur32.dd
'
Public Const NameSamCompatible As Long = 2
Public Declare Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As Long, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long

' advapi32.dll
'
Public Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameW" (ByVal lpSystemName As Long, ByVal lpAccountName As Long, ByVal Sid As Long, ByRef cbSid As Long, ByVal ReferencedDomainName As Long, ByRef cbReferencedDomainName As Long, ByRef peUse As Long) As Long

' user32.dll
'
Public Declare Function GetProcessWindowStation Lib "user32.dll" () As Long
Public Declare Function GetUserObjectInformation Lib "user32.dll" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, ByRef pvInfo As Any, ByVal nLength As Long, ByRef lpnLengthNeeded As Long) As Long
Public Declare Function GetSystemMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long


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
    Dim Info As OSVERSIONINFOA
    Info.dwOSVersionInfoSize = Len(Info)

    If GetVersionExA(Info) = BOOL_FALSE Then _
        Throw Cor.NewInvalidOperationException("Could not load operating system information.")

    If Info.dwPlatformId = PlatformID.Win32NT Then
        Set API = New Win32ApiW
    Else
        Set API = New Win32ApiA
    End If
End Sub
