VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: ErrorHelper
'
Option Explicit

Friend Function Argument(ByVal Key As ResourceStringKey, Optional ByRef ParamName As String) As ArgumentException
    Set Argument = Cor.NewArgumentException(Environment.GetResourceString(Key), ParamName)
End Function

Friend Function ArgumentNull(ByRef ParamName As String, Optional ByVal Key As ResourceStringKey = ArgumentNull_Exception) As ArgumentNullException
    Set ArgumentNull = Cor.NewArgumentNullException(ParamName, Environment.GetResourceString(Key))
End Function

Friend Function ArgumentOutOfRange(ByRef ParamName As String, Optional ByVal Key As ResourceStringKey = ArgumentOutOfRange_Exception) As ArgumentOutOfRangeException
    Set ArgumentOutOfRange = Cor.NewArgumentOutOfRangeException(ParamName, Environment.GetResourceString(Key))
End Function

Friend Function Rank(Optional ByVal Key As ResourceStringKey = Rank_MultiDimNotSupported) As RankException
    Set Rank = Cor.NewRankException(Environment.GetResourceString(Key))
End Function

Friend Function ObjectDisposed(Optional ByVal Key As ResourceStringKey = ObjectDisposed_Generic) As ObjectDisposedException
    Set ObjectDisposed = Cor.NewObjectDisposedException("", Environment.GetResourceString(Key))
End Function

Friend Function IndexOutOfRange(ByVal Key As ResourceStringKey) As IndexOutOfRangeException
    Set IndexOutOfRange = Cor.NewIndexOutOfRangeException(Environment.GetResourceString(Key))
End Function

Friend Function InvalidOperation(ByVal Key As ResourceStringKey) As InvalidOperationException
    Set InvalidOperation = Cor.NewInvalidOperationException(Environment.GetResourceString(Key))
End Function

Friend Function NotSupported(ByVal Key As ResourceStringKey) As NotSupportedException
    Set NotSupported = Cor.NewNotSupportedException(Environment.GetResourceString(Key))
End Function

Friend Function IO(ByVal Key As ResourceStringKey) As IOException
    Set IO = Cor.NewIOException(Environment.GetResourceString(Key))
End Function

Friend Sub WinIOError(ByVal ErrorCode As Long, ByRef Path As String)
    Select Case ErrorCode
        Case ERROR_PATH_NOT_FOUND
            If Len(Path) = 0 Then _
                Throw Cor.NewDirectoryNotFoundException(Environment.GetResourceString(IO_PathNotFound_NoPathName))
            
            Throw Cor.NewDirectoryNotFoundException(Environment.GetResourceString(IO_PathNotFound_Path, Path))
            
        Case ERROR_FILE_NOT_FOUND
            If Len(Path) = 0 Then _
                Throw Cor.NewFileNotFoundException(Message:=Environment.GetResourceString(IO_FileNotFound))
            
            Throw Cor.NewFileNotFoundException(Path, Environment.GetResourceString(IO_FileNotFound_Name, Path))
            
        Case ERROR_ACCESS_DENIED
            If Len(Path) = 0 Then _
                Throw Cor.NewUnauthorizedAccessException(Environment.GetResourceString(UnauthorizedAccess_IODenied_NoPathName))
                
            Throw Cor.NewUnauthorizedAccessException(Environment.GetResourceString(UnauthorizedAccess_IODenied_Path, Path))
            
        Case ERROR_ALREADY_EXISTS
            If Len(Path) > 0 Then _
                Throw Cor.NewIOException(Environment.GetResourceString(IO_AlreadyExists_Name, Path), MakeHRFromErrorCode(ErrorCode))
            
        Case ERROR_FILENAME_EXCED_RANGE
            Throw Cor.NewPathTooLongException(Environment.GetResourceString(IO_PathTooLong))
            
        Case ERROR_INVALID_DRIVE
            Throw Cor.NewDriveNotFoundException(Environment.GetResourceString(IO_DriveNotFound_Drive, Path))
            
        Case ERROR_SHARING_VIOLATION
            If Len(Path) = 0 Then _
                Throw Cor.NewIOException(Environment.GetResourceString(IO_SharingViolation_NoFileName), MakeHRFromErrorCode(ErrorCode))
                
            Throw Cor.NewIOException(Environment.GetResourceString(IO_SharingViolation_File, Path), MakeHRFromErrorCode(ErrorCode))
            
        Case ERROR_FILE_EXISTS
            If Len(Path) > 0 Then _
                Throw Cor.NewIOException(Environment.GetResourceString(IO_FileExists_Name, Path), MakeHRFromErrorCode(ErrorCode))
            
    End Select
    
    Throw Cor.NewIOException(GetMessage(ErrorCode), MakeHRFromErrorCode(ErrorCode))
End Sub
