VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKeyStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2016 Kelly Ethridge
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
' Module: RegistryKeyStatic
'

''
' Provides static methods for the <b>RegistryKey</b> class.
'
' @remarks This class cannot be directly instantiated. To access the
' methods, use the variable name directly.
' <pre>
' Dim Key As RegistryKey
' Set Key = RegistryKey.OpenRemoteBaseKey(LocalMachine, "RemoteName")
' </pre>
'
' @see Registry
' @see PublicFunctions
'
Option Explicit


''
' The root registry key to open on the remote machine.
'
' @param ClassesRoot Specifies the HKEY_CLASSES_ROOT key.
' @param CurrentConfig Specifies the HKEY_CURRENT_CONFIG key.
' @param CurrentUser Specifies the HKEY_CURRENT_USER key.
' @param LocalMachine Specifies the HKEY_LOCAL_MACHINE key.
' @param Users Specifies The HKEY_USERS key.
' @param PerformanceData Specifies the HKEY_PERFORMANCE_DATA key.
'
Public Enum RegistryHive
    ClassesRoot = &H80000000
    CurrentConfig = &H80000005
    CurrentUser = &H80000001
    LocalMachine = &H80000002
    Users = &H80000003
    PerformanceData = &H80000004
End Enum


''
' Connects to a Registry on a remote machine and returns a <b>RegistryKey</b>
' to the specified registry key on that machine.
'
' @param hKey The registry key on the remote machine to open.
' @param MachineName The name of the remote machine.
' @return A <b>RegistryKey</b> connected to the registry of a remote machine.
' @remarks If the <i>MachineName</i> is an empty String, then the registry of
' the local machine is opened.
'
Public Function OpenRemoteBaseKey(ByVal hKey As RegistryHive, ByVal MachineName As String) As RegistryKey
    Dim Result As Long
    Dim NewKey As Long
    
    Result = RegConnectRegistry(MachineName, hKey, NewKey)
    If Result <> ERROR_SUCCESS Then
        Error.Win32Error Result
    End If
    
    Set OpenRemoteBaseKey = Cor.NewRegistryKey(NewKey, GetRootName(NewKey), True)
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetRootName(ByVal hKey As Long) As String
    Select Case hKey
        Case HKEY_CLASSES_ROOT
            GetRootName = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_CONFIG
            GetRootName = "HKEY_CURRENT_CONFIG"
        Case HKEY_CURRENT_USER
            GetRootName = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE
            GetRootName = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            GetRootName = "HKEY_USERS"
        Case HKEY_PERFORMANCE_DATA
            GetRootName = "HKEY_PERFORMANCE_DATA"
    End Select
End Function

