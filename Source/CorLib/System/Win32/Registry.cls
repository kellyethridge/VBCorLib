VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
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
'    Module: Registry
'

''
' Supplies the base Registrykeys that access values and subkeys in the registry.
'
' @remarks These are the base sets of root registry keys used to access different
' classes of registry keys.
' <p>To access these methods, use the <b>Registry.*</b> syntax:
' <pre>
'     Dim rk As RegistryKey
'     Set rk = Registry.ClassesRoot
' </pre>
' @see RegistryKey
' @see RegistryKeyStatic
'
Option Explicit

Private mClassesRoot        As RegistryKey
Private mCurrentConfig      As RegistryKey
Private mCurrentUser        As RegistryKey
Private mDynData            As RegistryKey
Private mLocalMachine       As RegistryKey
Private mPerformanceData    As RegistryKey
Private mUsers              As RegistryKey
Private mMaxValueNameLength As Long
Private mMaxValueLength     As Long



Public Property Get MaxValueNameLength() As Long
    MaxValueNameLength = mMaxValueNameLength
End Property

Public Property Get MaxValueLength() As Long
    MaxValueLength = mMaxValueLength
End Property

''
' Returns a <b>RegistryKey</b> object that is then used to access registry keys and
' values for class types and their properties. The root section of the Registry
' accessed is HKEY_CLASSES_ROOT.
'
' @return Returns a <b>Registrykey</b> that is used to access the Class Root keys.
'
Public Property Get ClassesRoot() As RegistryKey
    Set ClassesRoot = mClassesRoot
End Property

''
' Returns a <b>RegistryKey</b> object that is then used to access registry keys and
' values for the current configuration of hardware. This is not user specific data.
' The root section of the Registry accessed is HKEY_CURRENT_CONFIG.
'
Public Property Get CurrentConfig() As RegistryKey
    Set CurrentConfig = mCurrentConfig
End Property

''
' Returns a <b>RegistryKey</b> object that is then used to access registry keys and
' values for information and preferences of the current user. The root section of
' the Registry accessed is HKEY_CURRENT_USER.
'
Public Property Get CurrentUser() As RegistryKey
    Set CurrentUser = mCurrentUser
End Property

''
' Contains dynamic registry data. This field reads the Windows registry base key HKEY_DYN_DATA.
'
' @remarks This root class is only available of Windows 95/98/ME. An exception
' is thrown for other windows platforms.
'
Public Property Get DynData() As RegistryKey
    Set DynData = mDynData
End Property

''
' Returns a <b>RegistryKey</b> object that can be used to access information and preferences
' about the local user machine. The root section of the Registry accessed is HKEY_LOCAL_MACHINE.
'
Public Property Get LocalMachine() As RegistryKey
    Set LocalMachine = mLocalMachine
End Property

''
' Contains performance information for software components. This field reads the Windows registry
' base key HKEY_PERFORMANCE_DATA.
'
Public Property Get PerformanceData() As RegistryKey
    Set PerformanceData = mPerformanceData
End Property

''
' Contains information about the default user configuration. This field reads the Windows
' registry base key HKEY_USERS.
'
Public Property Get Users() As RegistryKey
    Set Users = mUsers
End Property

''
' Returns a value from the registry key.
'
' @param KeyName The full registry key name that contains the value.
' @param ValueName The name of the value in the registry key.
' @param DefaultValue The value to return if the registry value does not exist.
' @return A value from the registry.
' @remarks A full registry key name must include the root level key.
' <pre>
' v = Registry.GetValue("HKEY_CURRENT_USER\Environment\", "TEMP")
' </pre>
'
Public Function GetValue(ByVal KeyName As String, ByVal ValueName As String, Optional ByRef DefaultValue As Variant) As Variant
    Dim rk As RegistryKey
    
    Set rk = GetSubKey(KeyName)
    If Not rk Is Nothing Then Call Helper.MoveVariant(GetValue, rk.GetValue(ValueName, DefaultValue))
End Function

''
' Sets a value to a registry key, creating it if it does not exist.
'
' @param KeyName The full registry key name to write the value to.
' @param ValueName The name of the value to be written in the key.
' @param Value The value to be written in the key.
' @remarks A full registry key name must include the root level key.
' <pre>
' Registry.SetValue "HKEY_CURRENT_USER\Environment\", "MyValue", "Hello"
' </pre>
'
Public Sub SetValue(ByVal KeyName As String, ByVal ValueName As String, ByRef Value As Variant)
    Dim rk As RegistryKey
    
    Set rk = GetSubKey(KeyName, True, True)
    If rk Is Nothing Then _
        Throw Cor.NewArgumentException("Could not save to registry key", "KeyName")
    
    Call rk.SetValue(ValueName, Value)
End Sub

''
' Returns a root RegistryKey for the key defined in the key name.
'
' @param KeyName The keyname path, including the root name at the beginning.
' @return The root RegistryKey.
' @remarks The KeyName can contain subkeys. The root key name will be parsed out.
'
Public Function GetRootKey(ByVal KeyName As String) As RegistryKey
    If Len(KeyName) = 0 Then Exit Function
    KeyName = cString.Trim(KeyName, "\")
    
    Dim Keys() As String
    Keys = Split(KeyName, "\")
    
    Select Case UCase$(Keys(0))
        Case "HKEY_CLASSES_ROOT":       Set GetRootKey = Me.ClassesRoot
        Case "HKEY_CURRENT_CONFIG":     Set GetRootKey = Me.CurrentConfig
        Case "HKEY_CURRENT_USER":       Set GetRootKey = Me.CurrentUser
        Case "HKEY_DYN_DATA":           Set GetRootKey = Me.DynData
        Case "HKEY_LOCAL_MACHINE":      Set GetRootKey = Me.LocalMachine
        Case "HKEY_PERFORMANCE_DATA":   Set GetRootKey = Me.PerformanceData
        Case "HKEY_USERS":              Set GetRootKey = Me.Users
    End Select
End Function

''
' Returns a RegistryKey to the full path of the key name.
'
' @param KeyName The full registry key path, including the root name.
' @param AutoCreate Flag to indicate if the full path should be created if it does not exist.
' @param Writable Flag to indicate if an existing path is writable.
' @return A RegistryKey to the full path of the key name, or Nothing if the path does not exist.
' @remarks If AutoCreate is true, then Writable is ignored.
'
Public Function GetSubKey(ByVal KeyName As String, Optional ByVal AutoCreate As Boolean = False, Optional ByVal Writable As Boolean = False) As RegistryKey
    If Len(KeyName) = 0 Then Exit Function
    
    Dim Ret As RegistryKey
    Set Ret = GetRootKey(KeyName)
    If Ret Is Nothing Then Exit Function
    
    Dim i As Long
    i = InStr(KeyName, "\")
    If i = 0 Then Exit Function
        
    If AutoCreate Then
        Set GetSubKey = Ret.CreateSubKey(Mid$(KeyName, i))
    Else
        Set GetSubKey = Ret.OpenSubKey(Mid$(KeyName, i), Writable)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    If Environment.IsNT Then
        mMaxValueNameLength = 16383
        mMaxValueLength = -1
    Else
        mMaxValueNameLength = 255
        mMaxValueLength = 16300
    End If
    
    Set mClassesRoot = Cor.NewRegistryKey(HKEY_CLASSES_ROOT, "HKEY_CLASSES_ROOT", True)
    Set mCurrentConfig = Cor.NewRegistryKey(HKEY_CURRENT_CONFIG, "HKEY_CURRENT_CONFIG", True)
    Set mCurrentUser = Cor.NewRegistryKey(HKEY_CURRENT_USER, "HKEY_CURRENT_USER", True)
    Set mDynData = Cor.NewRegistryKey(HKEY_DYN_DATA, "HKEY_DYN_DATA", True)
    Set mLocalMachine = Cor.NewRegistryKey(HKEY_LOCAL_MACHINE, "HKEY_LOCAL_MACHINE", True)
    Set mPerformanceData = Cor.NewRegistryKey(HKEY_PERFORMANCE_DATA, "HKEY_PERFORMANCE_DATA", True)
    Set mUsers = Cor.NewRegistryKey(HKEY_USERS, "HKEY_USERS", True)
End Sub
