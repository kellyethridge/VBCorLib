Attribute VB_Name = "IOConstants"
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
' Module: IOConstants
'
Option Explicit

Public Const vbPathSeparatorChar            As Integer = vbSemiColonChar
Public Const vbVolumeSeparatorChar          As Integer = vbColonChar
Public Const vbDirectorySeparatorChar       As Integer = vbBackslashChar
Public Const vbAltDirectorySeparatorChar    As Integer = vbForwardSlashChar
Public Const vbDoubleDirectorySeparatorChar As Long = &H5C005C
Public Const vbPathSeparator                As String = ";"
Public Const vbVolumeSeparator              As String = ":"
Public Const vbDirectorySeparator           As String = "\"
Public Const vbAltDirectorySeparator        As String = "/"

Public Const MAX_PATH                   As Long = 260
Public Const MAX_DIR                    As Long = 255
Public Const MAX_LONG_PATH              As Long = 32000

Public Const FILE_FLAG_OVERLAPPED       As Long = &H40000000
Public Const FILE_ATTRIBUTE_NORMAL      As Long = &H80
Public Const INVALID_HANDLE_VALUE       As Long = -1
Public Const FILE_TYPE_DISK             As Long = &H1
Public Const FILE_ATTRIBUTE_DIRECTORY   As Long = &H10
Public Const INVALID_FILE_ATTRIBUTES    As Long = -1
Public Const INVALID_SET_FILE_POINTER  As Long = -1
Public Const INVALID_FILE_SIZE         As Long = -1
Public Const ERROR_BROKEN_PIPE         As Long = 109

' File manipulation function attributes
Public Const GENERIC_READ               As Long = &H80000000
Public Const GENERIC_WRITE              As Long = &H40000000
Public Const OPEN_EXISTING              As Long = 3
Public Const PAGE_READONLY              As Long = &H2
Public Const PAGE_READWRITE             As Long = &H4
Public Const PAGE_WRITECOPY             As Long = &H8
Public Const INVALID_HANDLE             As Long = -1
Public Const FILE_SHARE_READ            As Long = 1
Public Const FILE_SHARE_WRITE           As Long = 2
Public Const STANDARD_RIGHTS_REQUIRED   As Long = &HF0000
Public Const SECTION_QUERY              As Long = &H1
Public Const SECTION_MAP_WRITE          As Long = &H2
Public Const SECTION_MAP_READ           As Long = &H4
Public Const SECTION_MAP_EXECUTE        As Long = &H8
Public Const SECTION_EXTEND_SIZE        As Long = &H10
Public Const SECTION_ALL_ACCESS         As Long = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Public Const FILE_MAP_READ              As Long = SECTION_MAP_READ
Public Const FILE_MAP_ALL_ACCESS        As Long = SECTION_ALL_ACCESS
Public Const NULL_HANDLE                As Long = 0

Public Const NO_ERROR                   As Long = 0
Public Const ERROR_PATH_NOT_FOUND       As Long = 3
Public Const ERROR_ACCESS_DENIED        As Long = 5
Public Const ERROR_FILE_NOT_FOUND       As Long = 2
Public Const ERROR_FILE_EXISTS          As Long = 80
Public Const ERROR_INSUFFICIENT_BUFFER  As Long = 122
