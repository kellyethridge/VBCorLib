Attribute VB_Name = "RegistryConstants"
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: RegistryConstants
'
Option Explicit

' Registry Constants
Public Const REG_NONE                      As Long = 0
Public Const REG_UNKNOWN                   As Long = 0
Public Const REG_SZ                        As Long = 1
Public Const REG_DWORD                     As Long = 4
Public Const REG_BINARY                    As Long = 3
Public Const REG_MULTI_SZ                  As Long = 7
Public Const REG_EXPAND_SZ                 As Long = 2
Public Const REG_QWORD                     As Long = 11

Public Const ERROR_SUCCESS                 As Long = 0
Public Const ERROR_INVALID_HANDLE          As Long = 6
Public Const ERROR_INVALID_PARAMETER       As Long = 87
Public Const ERROR_CALL_NOT_IMPLEMENTED    As Long = 120
Public Const ERROR_MORE_DATA               As Long = 234
Public Const ERROR_NO_MORE_ITEMS           As Long = 259
Public Const ERROR_CANTOPEN                As Long = 1011
Public Const ERROR_CANTREAD                As Long = 1012
Public Const ERROR_CANTWRITE               As Long = 1013
Public Const ERROR_REGISTRY_RECOVERED      As Long = 1014
Public Const ERROR_REGISTRY_CORRUPT        As Long = 1015
Public Const ERROR_REGISTRY_IO_FAILED      As Long = 1016
Public Const ERROR_NOT_REGISTRY_FILE       As Long = 1017
Public Const ERROR_KEY_DELETED             As Long = 1018

' Registry Root Keys
Public Const HKEY_CLASSES_ROOT      As Long = &H80000000
Public Const HKEY_CURRENT_CONFIG    As Long = &H80000005
Public Const HKEY_CURRENT_USER      As Long = &H80000001
Public Const HKEY_DYN_DATA          As Long = &H80000006
Public Const HKEY_LOCAL_MACHINE     As Long = &H80000002
Public Const HKEY_USERS             As Long = &H80000003
Public Const HKEY_PERFORMANCE_DATA  As Long = &H80000004

' Registry Flags
Public Const READ_CONTROL           As Long = &H20000
Public Const STANDARD_RIGHTS_ALL    As Long = &H1F0000
Public Const STANDARD_RIGHTS_READ   As Long = READ_CONTROL
Public Const KEY_QUERY_VALUE        As Long = &H1
Public Const KEY_SET_VALUE          As Long = &H2
Public Const KEY_CREATE_SUB_KEY     As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_CREATE_LINK        As Long = &H20
Public Const KEY_NOTIFY             As Long = &H10
Public Const SYNCHRONIZE            As Long = &H100000
Public Const KEY_READ               As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS         As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))



