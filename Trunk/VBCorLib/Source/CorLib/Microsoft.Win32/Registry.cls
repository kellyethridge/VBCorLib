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
' Module: Registry
'

''
' Supplies the base Registrykeys that access values and subkeys in the registry.
'
' @remarks These are the base sets of root registry keys used to access different
' classes of registry keys.
' <p>To access these methods, use the <b>Registry.*</b> syntax:
' <pre>
'     Dim Key As RegistryKey
'     Set Key = Registry.ClassesRoot
' </pre>
' @see RegistryKey
' @see RegistryKeyStatic
'
Option Explicit

Private mClassesRoot        As RegistryKey
Private mCurrentConfig      As RegistryKey
Private mCurrentUser        As RegistryKey
Private mLocalMachine       As RegistryKey
Private mPerformanceData    As RegistryKey
Private mUsers              As RegistryKey



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
Public Function GetValue(ByRef KeyName As String, ByRef ValueName As String, Optional ByRef DefaultValue As Variant) As Variant
    Dim Key As RegistryKey
    
    Set Key = GetSubKey(KeyName, False)
    Helper.MoveVariant GetValue, Key.GetValue(ValueName, DefaultValue)
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
Public Sub SetValue(ByRef KeyName As String, ByRef ValueName As String, ByRef Value As Variant)
    Dim Key As RegistryKey
    
    Set Key = GetSubKey(KeyName, True)
    Key.SetValue ValueName, Value
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetRootKey(ByRef KeyName As String) As RegistryKey
    Select Case UCase$(KeyName)
        Case "HKEY_CLASSES_ROOT"
            Set GetRootKey = Me.ClassesRoot
        Case "HKEY_CURRENT_CONFIG"
            Set GetRootKey = Me.CurrentConfig
        Case "HKEY_CURRENT_USER"
            Set GetRootKey = Me.CurrentUser
        Case "HKEY_LOCAL_MACHINE"
            Set GetRootKey = Me.LocalMachine
        Case "HKEY_PERFORMANCE_DATA"
            Set GetRootKey = Me.PerformanceData
        Case "HKEY_USERS"
            Set GetRootKey = Me.Users
        Case Else
            Error.Argument Arg_RegInvalidKeyName
    End Select
End Function

Private Function GetSubKey(ByRef KeyName As String, ByVal Writable As Boolean) As RegistryKey
    Dim BaseKeyName As String
    Dim Index       As Long
    
    Index = InStr(KeyName, "\")
    If Index > 0 Then
        BaseKeyName = Left$(KeyName, Index - 1)
    End If
    
    Dim Root As RegistryKey
    Set Root = GetRootKey(BaseKeyName)
    
    If Writable Then
        Set GetSubKey = Root.CreateSubKey(Mid$(KeyName, Index))
    Else
        Set GetSubKey = Root.OpenSubKey(Mid$(KeyName, Index))
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mClassesRoot = Cor.NewRegistryKey(HKEY_CLASSES_ROOT, "HKEY_CLASSES_ROOT", True)
    Set mCurrentConfig = Cor.NewRegistryKey(HKEY_CURRENT_CONFIG, "HKEY_CURRENT_CONFIG", True)
    Set mCurrentUser = Cor.NewRegistryKey(HKEY_CURRENT_USER, "HKEY_CURRENT_USER", True)
    Set mLocalMachine = Cor.NewRegistryKey(HKEY_LOCAL_MACHINE, "HKEY_LOCAL_MACHINE", True)
    Set mPerformanceData = Cor.NewRegistryKey(HKEY_PERFORMANCE_DATA, "HKEY_PERFORMANCE_DATA", True)
    Set mUsers = Cor.NewRegistryKey(HKEY_USERS, "HKEY_USERS", True)
End Sub
