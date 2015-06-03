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
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: RegistryKeyStatic
'

''
' Provides static methods for the <b>RegistryKey</b> class.
'
' @remarks This class cannot be directly instantiated. To access the
' methods, use the variable name directly.
' <pre>
' Dim rk As RegistryKey
' Set rk = RegistryKey.OpenRemoteBaseKey(LocalMachine, "RemoteName")
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
' @param DynData Specifies the HKEY_DYN_DATA key.
' @param LocalMachine Specifies the HKEY_LOCAL_MACHINE key.
' @param Users Specifies The HKEY_USERS key.
' @param PerformanceData Specifies the HKEY_PERFORMANCE_DATA key.
'
Public Enum RegistryHive
    ClassesRoot = &H80000000
    CurrentConfig = &H80000005
    CurrentUser = &H80000001
    DynData = &H80000006
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
    If Result <> ERROR_SUCCESS Then Error.ApiError Result
    
    Set OpenRemoteBaseKey = Cor.NewRegistryKey(NewKey, GetRootName(NewKey), True)
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetRootName(ByVal hKey As Long) As String
    Dim Ret As String
    
    Select Case hKey
        Case HKEY_CLASSES_ROOT:         Ret = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_CONFIG:       Ret = "HKEY_CURRENT_CONFIG"
        Case HKEY_CURRENT_USER:         Ret = "HKEY_CURRENT_USER"
        Case HKEY_DYN_DATA:             Ret = "HKEY_DYN_DATA"
        Case HKEY_LOCAL_MACHINE:        Ret = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS:                Ret = "HKEY_USERS"
        Case HKEY_PERFORMANCE_DATA:     Ret = "HKEY_PERFORMANCE_DATA"
    End Select
    
    GetRootName = Ret
End Function

