VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: Directory
'

''
' Provides a set of static methods for manipulating and retrieving directory information.
'
Option Explicit

Private Const FILE_FLAG_BACKUP_SEMANTICS        As Long = &H2000000
Private Const ERROR_NO_MORE_FILES               As Long = 18&
Private Const DOUBLE_PERIOD                     As Long = &H2E002E

Private Enum SystemEntryTypes
    FileEntry = 1
    DirectoryEntry = 2
End Enum

Public Enum SearchOption
    AllDirectories
    TopDirectoryOnly
End Enum

Private mDirectorySeparatorChar      As Integer
Private mAltDirectorySeparatorChar   As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mDirectorySeparatorChar = Path.DirectorySeparatorChar
    mAltDirectorySeparatorChar = Path.AltDirectorySeparatorChar
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   API
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
' Creates a directory and any subdirectories in the specified path.
'
' @param Path The directory path to be created.
' @return A DirectoryInfo object that represents the created directory.
' @remarks The path to be created can be a relative or absolute path.
'
Public Function CreateDirectory(ByVal Path As String) As DirectoryInfo
    If Len(Path) = 0 Then _
        Throw Error.Argument(Argument_EmptyPath, "Path")
    
    Dim FullPath As String
    FullPath = Statics.Path.GetFullPath(Path)
    
    If Len(FullPath) > 1 And Statics.Path.IsDirectorySeparator(CorString.LastChar(FullPath)) Then
        FullPath = Left$(FullPath, Len(FullPath) - 1)
    End If
    
'    Statics.Path.FixupPath Path
'    Path = CorString.TrimEnd(Path, Statics.Path.DirectorySeparator)
    
'    If Len(Path) > Statics.Path.MaxPath Then _
'        Throw New PathTooLongException
'    If InStr(3, Path, Statics.Path.VolumeSeparator) > 0 Then _
'        Throw New NotSupportedException
    
    Dim DirectoryNames() As String
    DirectoryNames = Split(FullPath, Statics.Path.DirectorySeparator)
    
    Dim i As Long
    i = CorArray.GetLength(DirectoryNames)
    Do
        Dim TestDirectory As String
        TestDirectory = CorString.Join(Statics.Path.DirectorySeparator, DirectoryNames, 0, i)
        If Directory.Exists(TestDirectory) Then
            Exit Do
        End If
        
        i = i - 1
    Loop While i > 0

    Do While i <= UBound(DirectoryNames)
        i = i + 1
        Dim NewDirectory As String
        NewDirectory = CorString.Join(Statics.Path.DirectorySeparator, DirectoryNames, 0, i)
        If Api.CreateDirectory(NewDirectory) = BOOL_FALSE Then
            IOError Err.LastDllError
        End If
    Loop
    
    Set CreateDirectory = Cor.NewDirectoryInfo(FullPath)
End Function

''
' Checks if the path exists.
'
' @param Path The path to check for existence.
' @return Indication of the existence of the path.
' @remarks The path can be relative or absolute.
'
Public Function Exists(ByVal Path As String) As Boolean
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    
    If File.GetFileData(Path, Data) = NO_ERROR Then
        Exists = CBool(Data.dwFileAttributes And FileAttributes.DirectoryAttr)
    End If
End Function

''
' Deletes directories and files within directories.
'
' @param Path The top directory to be deleted.
' @param Recursive If set to True, then all directories and files in the top
' directory will be deleted along with the top directory. Otherwise, the specified
' directory must be empty to be deleted.
'
Public Sub Delete(ByRef Path As String, Optional ByVal Recursive As Boolean = False)
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath))
    
    Dim NormalizedPath As String
    NormalizedPath = CorString.TrimEnd(Statics.Path.GetFullPath(Path), "/\")
    
    If Recursive Then
        Dim Data As WIN32_FIND_DATA
        
        With SafeFindFirstFile(NormalizedPath & "\*", Data)
            If Not .IsInvalid Then
                ' Simply loop through any files or directories
                ' within this directory and delete them.
                Do
                    Dim FileName As String
                    FileName = GetFileNameFromFindData(Data)
                    
                    If Len(FileName) > 0 Then
                        Dim FullPath As String
                        FullPath = Statics.Path.Combine(NormalizedPath, FileName)
                        
                        If CBool(Data.dwFileAttributes And FileAttributes.DirectoryAttr) Then
                            Delete FullPath, True
                        Else
                            If Api.DeleteFile(FullPath) = BOOL_FALSE Then IOError Err.LastDllError, FullPath
                        End If
                    End If
                Loop While Api.FindNextFile(.DangerousGetHandle, Data) = BOOL_TRUE
                
                Dim Result As Long
                Result = Err.LastDllError
                
                If Result <> ERROR_NO_MORE_FILES Then IOError Result, FullPath
            End If
        End With
    End If
    
    If Api.RemoveDirectory(NormalizedPath) = BOOL_FALSE Then
        If Err.LastDllError <> ERROR_PATH_NOT_FOUND Then IOError Err.LastDllError, Path
    End If
End Sub

''
' Returns the time of creation for the specified directory in local time.
'
' @param Path The directory to retrieve the creation time of.
' @return A CorDateTime object containing the creation time of the directory.
' @remarks The path can be relative or absolute.
'
Public Function GetCreationTime(ByVal Path As String) As CorDateTime
    Set GetCreationTime = File.GetCreationTime(Path)
End Function

''
' Returns the time of creation for the specified directory in UTC time.
'
' @param Path The directory to retrieve the creation time of.
' @return A CorDateTime object containing the creation time of the directory.
' @remarks The path can be relative or absolute.
'
Public Function GetCreationTimeUtc(ByVal Path As String) As CorDateTime
    Set GetCreationTimeUtc = File.GetCreationTimeUtc(Path)
End Function

''
' Returns the time the directory was last accessed in local time.
'
' @param Path The directory to retrieve the last time it was accessed.
' @return A CorDateTime object containing the time the directory was last accessed.
' @remarks The path can be relative or absolute.
'
Public Function GetLastAccessTime(ByVal Path As String) As CorDateTime
    Set GetLastAccessTime = File.GetLastAccessTime(Path)
End Function

''
' Returns the time the directory was last accessed in UTC time.
'
' @param Path The directory to retrieve the last time it was accessed.
' @return A CorDateTime object containing the time the directory was last accessed.
' @remarks The path can be relative or absolute.
'
Public Function GetLastAccessTimeUtc(ByVal Path As String) As CorDateTime
    Set GetLastAccessTimeUtc = File.GetLastAccessTimeUtc(Path)
End Function

''
' Returns the time the directory was last written to in local time.
'
' @param Path The directory to retrieve the last time it was written to.
' @return A CorDateTime object containing the last time the directory was written to.
' @remarks The path can relative or absolute.
'
Public Function GetLastWriteTime(ByVal Path As String) As CorDateTime
    Set GetLastWriteTime = File.GetLastWriteTime(Path)
End Function

''
' Returns the time the directory was last written to in UTC time.
'
' @param Path The directory to retrieve the last time it was written to.
' @return A CorDateTime object containing the last time the directory was written to.
' @remarks The path can relative or absolute.
'
Public Function GetLastWriteTimeUtc(ByVal Path As String) As CorDateTime
    Set GetLastWriteTimeUtc = File.GetLastWriteTimeUtc(Path)
End Function

''
' Returns the current directory the application is set to.
'
' @return The current set directory.
'
Public Function GetCurrentDirectory() As String
    Dim Result As String
    Result = String$(Path.MaxPath, 0)
    
    Dim Size As Long
    Size = Api.GetCurrentDirectory(Len(Result), Result)
    If Size = 0 Then IOError Err.LastDllError
    
    GetCurrentDirectory = Left$(Result, Size)
End Function

''
' Returns a list of all the directories found in the specified directory
' that matches the search pattern.
'
' @param Path The directory to find the requested directories.
' @param SearchPattern A pattern to compare all the directories against, returning
' those that match the pattern.
' @return An array of the matching directories. If no directories match the pattern, then
' an empty zero-length array.
' @remarks The path can be absolute or relative.
'
Public Function GetDirectories(ByVal Path As String, Optional ByVal SearchPattern As String = "*", Optional ByVal SearchOpt As SearchOption = TopDirectoryOnly) As String()
    GetDirectories = InternalGetFileSystemEntries(Path, SearchPattern, DirectoryEntry, SearchOpt)
End Function

''
' Returns the root directory of the specified path.
'
' @param Path The path to return the root of.
' @return The root of the specified directory.
' @remarks If the path is relative, then the current directory will
' be used as the root.
'
Public Function GetDirectoryRoot(ByVal Path As String) As String
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), Environment.GetResourceString(Parameter_Path))
    
    If Not Statics.Path.IsPathRooted(Path) Then
        Path = Statics.Path.GetFullPath(Path)
    End If
    
    GetDirectoryRoot = Statics.Path.GetPathRoot(Path)
End Function

''
' Returns a list of all the files that match a pattern within the specified directory.
'
' @param Path The directory to search for the matching files.
' @param SearchPattern A pattern to match the files against.
' @return The list of matching file names. If no filenames matched, then an empty,
' zero-length array is returned.
' @remarks The path can be absolute or relative. The path is not included in the
' filename result.
'
Public Function GetFiles(ByVal Path As String, Optional ByVal SearchPattern As String = "*", Optional ByVal SearchOpt As SearchOption = TopDirectoryOnly) As String()
    GetFiles = InternalGetFileSystemEntries(Path, SearchPattern, FileEntry, SearchOpt)
End Function

''
' Returns both directories and files that match the specified pattern.
'
' @param Path The directory to search for the files and directories.
' @param SearchPattern The pattern the filenames and directories must match.
' @return The list of files and directories that matched the pattern. If no files or directories
' matched the pattern, then an empty, zero-length array is returned.
' @remarks The path can be absolute or relative.
'
Public Function GetFileSystemEntries(ByVal Path As String, Optional ByVal SearchPattern As String = "*", Optional SearchOpt As SearchOption = TopDirectoryOnly) As String()
    GetFileSystemEntries = InternalGetFileSystemEntries(Path, SearchPattern, DirectoryEntry Or FileEntry, SearchOpt)
End Function

''
' Returns a list of the logical drives installed on this machine.
'
' @return The list of logical drives on this machine.
'
Public Function GetLogicalDrives() As String()
    Dim Drives() As String
    ReDim Drives(25) As String ' just allocate the maximum letters possible.
    
    Dim Mask As Long
    Mask = VBCorType.GetLogicalDrives
    
    Dim Letter As Long
    Letter = vbUpperAChar
    
    Dim i As Long
    Do While Mask <> 0
        If Mask And 1 Then
            Drives(i) = Chr$(Letter) & ":\"
            i = i + 1
        End If
        Letter = Letter + 1
        Mask = Mask \ 2
    Loop
    
    ReDim Preserve Drives(0 To i - 1)
    GetLogicalDrives = Drives
End Function

''
' Returns the parent of the specified directory.
'
' @param Path The directory to retrieve the parent of.
' @return A DirectoryInfo object representing the parent directory.
' @remarks The path can be absolute or relative.
'
Public Function GetParent(ByRef Path As String) As DirectoryInfo
    If Len(Path) = 0 Then _
        Throw Error.Argument(Argument_EmptyPath, "Path")

    Dim FullPath As String
    FullPath = Statics.Path.GetFullPath(Path)
    
    Dim DirName As String
    DirName = Statics.Path.GetDirectoryName(FullPath)
    If Len(DirName) > 0 Then
        Set GetParent = Cor.NewDirectoryInfo(DirName)
    End If
End Function

''
' Moves a directory and all of its contents to another location.
'
' @param SourceDirName The directory to be moved.
' @param DestDirName The directory to be moved to.
' @remarks The destination directory is the source directory renamed
' after the moving of the source directory.
'
Public Sub Move(ByRef SourceDirName As String, ByRef DestDirName As String)
    If Len(SourceDirName) = 0 Then _
        Throw Error.Argument(Argument_EmptyPath, "SourceDirName")
    If Len(DestDirName) = 0 Then _
        Throw Error.Argument(Argument_EmptyPath, "DestDirName")
    
    Dim FullSourceDirName   As String
    Dim FullDestDirName     As String
    
    FullSourceDirName = Path.GetFullPath(SourceDirName)
    FullDestDirName = Path.GetFullPath(DestDirName)
    
    If CorString.Equals(FullSourceDirName, FullDestDirName, OrdinalIgnoreCase) Then _
        Throw Cor.NewIOException("Source and Destination directories must be different.")
    If Path.GetPathRoot(FullSourceDirName) <> Path.GetPathRoot(FullDestDirName) Then _
        Throw Cor.NewIOException("Source and Destination directories must be on the same root drive.")
        
    If Api.MoveFile(FullSourceDirName, FullDestDirName) = BOOL_FALSE Then
        Dim Result  As Long
        Result = Err.LastDllError
        If Result = ERROR_FILE_NOT_FOUND Then
            IOError ERROR_PATH_NOT_FOUND, SourceDirName
        Else
            IOError Result
        End If
    End If
End Sub

''
' Changes the current directory to the specified directory.
'
' @param Path The new directory to set as the current directory.
' @remarks The path can be absolute or relative.
'
Public Sub SetCurrentDirectory(ByRef Path As String)
    If Len(Path) = 0 Then _
        Throw Error.Argument(Argument_EmptyPath, "Path")
    
'    Statics.Path.VerifyPath Path
    Dim FullPath As String
    FullPath = Statics.Path.GetFullPath(Path)

'    If Not Exists(Path) Then _
'        Throw New DirectoryNotFoundException
    
'    If Api.SetCurrentDirectory(Path) = BOOL_FALSE Then IOError Err.LastDllError, Path
    If Not Win32Native.SetCurrentDirectory(FullPath) Then
        Dim ErrorCode As Long
        ErrorCode = Err.LastDllError
        If ErrorCode = IOConstants.ERROR_FILE_NOT_FOUND Then
            ErrorCode = IOConstants.ERROR_PATH_NOT_FOUND
        End If
        IOError ErrorCode, FullPath
    End If
End Sub

''
' Sets the creation time of the specified directory to the local time.
'
' @param Path The directory to set the creation time for.
' @param CreationTime A Date or CorDateTime object of the new time.
' @remarks The time will have the UTC offset removed before setting the
' the directory to it. When the time is read, it will be in local time to the timezone.
'
Public Sub SetCreationTime(ByRef Path As String, ByRef CreationTime As Variant)
    SetCreationTimeUtc Path, TimeZone.CurrentTimeZone.ToUniversalTime(CreationTime)
End Sub

''
' Sets the creation time of the specified directory to UTC time.
'
' @param Path The directory to set the creation time for in UTC.
' @param CreationTime A Date or CorDateTime object of the new UTC time.
' @remarks The directory is set to the actual time passed in.
'
Public Sub SetCreationTimeUtc(ByRef Path As String, ByRef CreationTime As Variant)
    Dim Time As Currency
    Time = GetFileTimeUtc(CreationTime)
    SetFileTimeCore Path, VarPtr(Time), vbNullPtr, vbNullPtr
End Sub

''
' Sets the time the directory was last accessed in local time.
'
' @param Path The directory to set the new time for.
' @param LastAccessTime A Date or CorDateTime object of the new time.
' @remarks The UTC offset is removed from the time before setting the directory.
'
Public Sub SetLastAccessTime(ByRef Path As String, ByRef LastAccessTime As Variant)
    SetLastAccessTimeUtc Path, TimeZone.CurrentTimeZone.ToUniversalTime(LastAccessTime)
End Sub

''
' Sets the time the directory was last accessed in UTC time.
'
' @param Path The directory to set the new time for.
' @param LastAccessTime A Date or CorDateTime object of the new time.
' @remarks The directory is set to the actual time passed in.
'
Public Sub SetLastAccessTimeUtc(ByRef Path As String, ByRef LastAccessTime As Variant)
    Dim Time As Currency
    Time = GetFileTimeUtc(LastAccessTime)
    SetFileTimeCore Path, vbNullPtr, VarPtr(Time), vbNullPtr
End Sub

''
' Sets the time the directory was last written to in local time.
'
' @param Path The directory to set the new time for.
' @param LastWriteTime A Date or CorDateTime object of the new time.
'
Public Sub SetLastWriteTime(ByRef Path As String, ByRef LastWriteTime As Variant)
    SetLastWriteTimeUtc Path, TimeZone.CurrentTimeZone.ToUniversalTime(LastWriteTime)
End Sub

''
' Sets the time the directory was last written to in UTC time.
'
' @param Path The directory to set the new time for.
' @param LastWriteTime A Date or CorDateTime object of the new time.
'
Public Sub SetLastWriteTimeUtc(ByRef Path As String, ByRef LastWriteTime As Variant)
    Dim Time As Currency
    Time = GetFileTimeUtc(LastWriteTime)
    SetFileTimeCore Path, vbNullPtr, vbNullPtr, VarPtr(Time)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OpenHandle(ByRef Path As String) As SafeFileHandle
    Dim FullPath As String
    FullPath = Statics.Path.GetFullPath(Path)
    
    Dim root As String
    root = Statics.Path.GetPathRoot(FullPath)
    
    If (FullPath = root) And (CorString.CharAt(Path, 1) = vbVolumeSeparatorChar) Then _
        Throw Error.Argument(Arg_PathIsAVolume)
    
    Dim Handle As SafeFileHandle
    Set Handle = Win32Native.SafeCreateFile(FullPath, WriteAccess, WriteShare, OpenExisting, FILE_FLAG_BACKUP_SEMANTICS)
    If Handle.IsInvalid Then _
        IOError Err.LastDllError, FullPath
        
    Set OpenHandle = Handle
End Function

Private Function GetFileTimeUtc(ByRef Time As Variant) As Currency
    GetFileTimeUtc = CorDateTime.GetDateTime(Time).ToFileTimeUtc
End Function

Private Sub VerifySearchPattern(ByRef Pattern As String)
    Dim i As Long
    
    i = InStr(Pattern, "..")
    
    If i > 0 Then
        If i = Len(Pattern) - 1 Then _
            Throw Cor.NewArgumentException("Cannot end search pattern with '..'")
            
        Dim Ch As Integer
        Ch = MemWord(StrPtr(Pattern) + (i + 2) * 2)
        
        Select Case Ch
            Case mDirectorySeparatorChar, mAltDirectorySeparatorChar
                Throw Cor.NewArgumentException("'..' cannot be followed immediately by a directory separator.")
        End Select
    End If
End Sub

Private Function InternalGetFileSystemEntries(ByRef InitialPath As String, ByRef SearchPattern As String, ByVal EntriesToInclude As SystemEntryTypes, ByVal SearchOpt As SearchOption) As String()
    Dim FullPath As String
    FullPath = Path.GetFullPath(InitialPath)
    
    VerifySearchPattern SearchPattern
    
    Dim Entry       As WIN32_FIND_DATA
    Dim Entries     As New ArrayList
    Dim SubFolders  As New ArrayList
    Dim FileHandle  As Long
    
    With SafeFindFirstFile(Path.Combine(FullPath, SearchPattern), Entry)
        If Not .IsInvalid Then
            Do
                Dim EntryName As String
                EntryName = GetFileNameFromFindData(Entry)
                
                If Len(EntryName) > 0 Then
                    Dim NewPath As String
                    NewPath = Path.Combine(FullPath, EntryName)
                    
                    If EntryIsDirectory(Entry) Then
                        If CBool(EntriesToInclude And DirectoryEntry) Then
                            Entries.Add NewPath
                        End If
                        
                        If SearchOpt = AllDirectories Then
                            SubFolders.Add NewPath
                        End If
                    Else
                        If CBool(EntriesToInclude And FileEntry) Then
                            Entries.Add NewPath
                        End If
                    End If
                End If
            Loop While Api.FindNextFile(.DangerousGetHandle, Entry) <> BOOL_FALSE
        End If
        
        Dim Error As Long
        Error = Err.LastDllError
    
        If IsIOError(Error) Then IOError Error, InitialPath
    End With
    
    Dim SubFolder As Variant
    For Each SubFolder In SubFolders
        Entries.AddRange InternalGetFileSystemEntries(CStr(SubFolder), SearchPattern, EntriesToInclude, SearchOpt)
    Next

    InternalGetFileSystemEntries = Entries.ToArray(vbString)
End Function

Private Function EntryIsDirectory(ByRef Entry As WIN32_FIND_DATA) As Boolean
    EntryIsDirectory = CBool(Entry.dwFileAttributes And FileAttributes.DirectoryAttr)
End Function

Private Function IsIOError(ByVal Error As Long) As Boolean
    Select Case Error
        Case 0, ERROR_NO_MORE_FILES, ERROR_FILE_NOT_FOUND
        Case Else
            IsIOError = True
    End Select
End Function

Private Sub SetFileTimeCore(ByRef Path As String, ByVal CreationTimePtr As Long, ByVal LastAccessTimePtr As Long, ByVal LastWriteTimePtr As Long)
    With OpenHandle(Path)
        If SetFileTime(.DangerousGetHandle, ByVal CreationTimePtr, ByVal LastAccessTimePtr, ByVal LastWriteTimePtr) = BOOL_FALSE Then
            IOError Err.LastDllError, Path
        End If
    End With
End Sub

Private Function GetFileNameFromFindData(ByRef Data As WIN32_FIND_DATA) As String
    Dim Ret As String
    Ret = Data.cFileName
    If Len(Ret) = 0 Then
        Ret = Data.cAlternateFileName
    End If
    
    If (Data.dwFileAttributes And FileAttributes.DirectoryAttr) = FileAttributes.DirectoryAttr Then
        Select Case Len(Ret)
            Case 1: If Asc(Ret) = vbPeriodChar Then Exit Function
            Case 2: If AsLong(ByVal StrPtr(Ret)) = DOUBLE_PERIOD Then Exit Function
        End Select
    End If
    
    GetFileNameFromFindData = Ret
End Function

