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

Friend Function WinIOError(ByVal ErrCode As Long, ByRef Path As String) As Exception
    Dim Ex As IOException
    
    Select Case ErrCode
        Case ERROR_PATH_NOT_FOUND
            Set Ex = Cor.NewDirectoryNotFoundException("The directory '" & Path & "' could not be found.")
        Case ERROR_FILE_NOT_FOUND
            If Len(Path) = 0 Then
                Set Ex = Cor.NewFileNotFoundException(Message:=Environment.GetResourceString(IO_FileNotFound))
            Else
                Set Ex = Cor.NewFileNotFoundException(Path, Environment.GetResourceString(IO_FileNotFound_Name, Path))
            End If
        Case ERROR_ACCESS_DENIED
            Set Ex = Cor.NewInvalidOperationException("Permission to the specified file is denied.")
        Case Else
            Set Ex = Cor.NewIOException(GetMessage(ErrCode))
    End Select
    Ex.HResult = ErrCode
    
    Set WinIOError = Ex
End Function

'            switch (errorCode) {
'            Case Win32Native.ERROR_FILE_NOT_FOUND:
'                if (str.Length == 0)
'                    throw new FileNotFoundException(Environment.GetResourceString("IO.FileNotFound"));
'                Else
'                    throw new FileNotFoundException(Environment.GetResourceString("IO.FileNotFound_FileName", str), str);
'
'            Case Win32Native.ERROR_PATH_NOT_FOUND:
'                if (str.Length == 0)
'                    throw new DirectoryNotFoundException(Environment.GetResourceString("IO.PathNotFound_NoPathName"));
'                Else
'                    throw new DirectoryNotFoundException(Environment.GetResourceString("IO.PathNotFound_Path", str));
'
'            Case Win32Native.ERROR_ACCESS_DENIED:
'                if (str.Length == 0)
'                    throw new UnauthorizedAccessException(Environment.GetResourceString("UnauthorizedAccess_IODenied_NoPathName"));
'                Else
'                    throw new UnauthorizedAccessException(Environment.GetResourceString("UnauthorizedAccess_IODenied_Path", str));
'
'            Case Win32Native.ERROR_ALREADY_EXISTS:
'                if (str.Length == 0)
'                    goto default;
'                throw new IOException(Environment.GetResourceString("IO.IO_AlreadyExists_Name", str), Win32Native.MakeHRFromErrorCode(errorCode), maybeFullPath);
'
'            Case Win32Native.ERROR_FILENAME_EXCED_RANGE:
'                throw new PathTooLongException(Environment.GetResourceString("IO.PathTooLong"));
'
'            Case Win32Native.ERROR_INVALID_DRIVE:
'                throw new DriveNotFoundException(Environment.GetResourceString("IO.DriveNotFound_Drive", str));
'
'            Case Win32Native.ERROR_INVALID_PARAMETER:
'                throw new IOException(Win32Native.GetMessage(errorCode), Win32Native.MakeHRFromErrorCode(errorCode), maybeFullPath);
'
'            Case Win32Native.ERROR_SHARING_VIOLATION:
'                if (str.Length == 0)
'                    throw new IOException(Environment.GetResourceString("IO.IO_SharingViolation_NoFileName"), Win32Native.MakeHRFromErrorCode(errorCode), maybeFullPath);
'                Else
'                    throw new IOException(Environment.GetResourceString("IO.IO_SharingViolation_File", str), Win32Native.MakeHRFromErrorCode(errorCode), maybeFullPath);
'
'            Case Win32Native.ERROR_FILE_EXISTS:
'                if (str.Length == 0)
'                    goto default;
'                throw new IOException(Environment.GetResourceString("IO.IO_FileExists_Name", str), Win32Native.MakeHRFromErrorCode(errorCode), maybeFullPath);
'










