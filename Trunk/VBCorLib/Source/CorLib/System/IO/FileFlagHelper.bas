Attribute VB_Name = "FileFlagHelper"
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
' Module: FileFlagHelper
'
Option Explicit

Public Sub ValidateFileMode(ByVal Mode As FileMode)
    Select Case Mode
        Case FileMode.Append, FileMode.Create, FileMode.CreateNew, FileMode.OpenExisting, FileMode.OpenOrCreate, FileMode.Truncate
            Exit Sub
    End Select
    
    Throw Error.ArgumentOutOfRange("Mode", ArgumentOutOfRange_Enum)
End Sub

Public Sub ValidateFileAccess(ByVal Access As FileAccess)
    Select Case Access
        Case FileAccess.DefaultAccess, FileAccess.ReadAccess, FileAccess.ReadWriteAccess, FileAccess.WriteAccess
            Exit Sub
    End Select

    Throw Error.ArgumentOutOfRange("Access", ArgumentOutOfRange_Enum)
End Sub

Public Sub ValidateFileShare(ByVal Share As FileShare)
    Select Case Share
        Case FileShare.None, FileShare.ReadShare, FileShare.ReadWriteShare, FileShare.WriteShare
            Exit Sub
    End Select
    
    Throw Error.ArgumentOutOfRange("Share", ArgumentOutOfRange_Enum)
End Sub

Public Function GetFileModeDisplayName(ByVal Mode As FileMode) As String
    Dim Result As String
    
    Select Case Mode
        Case Append:        Result = "Append"
        Case Create:        Result = "Create"
        Case CreateNew:     Result = "CreateNew"
        Case OpenExisting:  Result = "OpenExisting"
        Case OpenOrCreate:  Result = "OpenOrCreate"
        Case Truncate:      Result = "Truncate"
        Case Else
            Throw Error.ArgumentOutOfRange("Mode", ArgumentOutOfRange_Enum)
    End Select
    
    GetFileModeDisplayName = Result
End Function

Public Function GetFileAccessDisplayName(ByVal Access As FileAccess) As String
    Dim Result As String
    
    Select Case Access
        Case DefaultAccess:     Result = "DefaultAccess"
        Case ReadAccess:        Result = "ReadAccess"
        Case WriteAccess:       Result = "WriteAccess"
        Case ReadWriteAccess:   Result = "ReadWriteAccess"
        Case Else
            Throw Error.ArgumentOutOfRange("Access", ArgumentOutOfRange_Enum)
    End Select
    
    GetFileAccessDisplayName = Result
End Function
