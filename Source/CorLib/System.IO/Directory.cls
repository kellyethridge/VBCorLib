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


''
' Creates a directory and any subdirectories in the specified path.
'
' @param Path The directory path to be created.
' @return A DirectoryInfo object that represents the created directory.
' @remarks The path to be created can be a relative or absolute path.
'
Public Function CreateDirectory(ByRef Path As String) As DirectoryInfo
    If Len(Path) = 0 Then _
        Error.Argument Argument_EmptyPath, "Path"
    
    Dim FullPath    As String
    Dim Paths       As New Stack
    
    FullPath = GetFullPath(Path)
    Set Paths = GetDirectoryPathsToCreate(FullPath)
    CreateDirectoryPaths Paths
    
    Set CreateDirectory = Cor.NewDirectoryInfo(FullPath)
End Function

Private Function GetFullPath(ByRef Path As String) As String
    Dim PathToExpand As String
    
    If Len(Path) > 1 And Right$(Path, 1) = Statics.Path.DirectorySeparator Then
        PathToExpand = Left$(Path, Len(Path) - 1)
    Else
        PathToExpand = Path
    End If
    
    GetFullPath = Statics.Path.GetFullPath(PathToExpand)
End Function

Private Function GetDirectoryPathsToCreate(ByRef FullPath As String) As Stack
    Dim PathParts() As String
    Dim NameIndex   As Long
    Dim Paths       As New Stack

    PathParts = Split(FullPath, Statics.Path.DirectorySeparator)
    NameIndex = CorArray.Length(PathParts)
    
    Do While NameIndex > 0
        Dim SubFolder As String
        SubFolder = CorString.Join(vbDirectorySeparator, PathParts, 0, NameIndex)
        
        If Directory.Exists(SubFolder) Then
            Exit Do
        End If
        
        Paths.Push SubFolder
        NameIndex = NameIndex - 1
    Loop

    Set GetDirectoryPathsToCreate = Paths
End Function

Private Sub CreateDirectoryPaths(ByVal Paths As Stack)
    Do While Paths.Count > 0
        If CreateDirectoryW(Paths.Pop, ByVal 0&) = BOOL_FALSE Then
            IOError.WinIOError Err.LastDllError
        End If
    Loop
End Sub

''
' Checks if the path exists.
'
' @param Path The path to check for existence.
' @return Indication of the existence of the path.
' @remarks The path can be relative or absolute.
'
Public Function Exists(ByRef Path As String) As Boolean
    Dim Data        As WIN32_FILE_ATTRIBUTE_DATA
    Dim FullPath    As String
    
    FullPath = Statics.Path.GetFullPath(Path)
    
    If File.GetFileData(FullPath, Data) = NO_ERROR Then
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
        Error.Argument Argument_EmptyPath
    
    Dim NormalizedPath As String
    NormalizedPath = CorString.TrimEnd(Statics.Path.GetFullPath(Path), "/\")
    
    If Recursive Then
        Dim Data As WIN32_FIND_DATAW
        
        With SafeFindFirstFile(NormalizedPath & "\*", Data)
            If Not .IsInvalid Then
                Do
                    Dim FileName As String
                    FileName = GetFileNameFromFindData(Data)
                    
                    If Len(FileName) > 0 Then
                        Dim FullPath As String
                        FullPath = Statics.Path.Combine(NormalizedPath, FileName)
                        
                        If PathIsDirectory(Data) Then
                            Delete FullPath, True
                        Else
                            If DeleteFileW(FullPath) = BOOL_FALSE Then
                                IOError.WinIOError Err.LastDllError, FullPath
                            End If
                        End If
                    End If
                Loop While FindNextFileW(.DangerousGetHandle, Data) = BOOL_TRUE
                
                Dim Result As Long
                Result = Err.LastDllError
                
                If Result <> ERROR_NO_MORE_FILES Then IOError.WinIOError Result, FullPath
            End If
        End With
    End If
    
    If RemoveDirectoryW(NormalizedPath) = BOOL_FALSE Then
        If Err.LastDllError <> ERROR_PATH_NOT_FOUND Then
            ThrowIOError Err.LastDllError, Path
        End If
    End If
End Sub

Private Function PathIsDirectory(ByRef Data As WIN32_FIND_DATAW) As Boolean
    PathIsDirectory = CBool(Data.dwFileAttributes And FileAttributes.DirectoryAttr)
End Function


''
' Returns the time of creation for the specified directory in local time.
'
' @param Path The directory to retrieve the creation time of.
' @return A CorDateTime object containing the creation time of the directory.
' @remarks The path can be relative or absolute.
'
Public Function GetCreationTime(ByRef Path As String) As CorDateTime
    Set GetCreationTime = File.GetCreationTime(Path)
End Function

''
' Returns the time of creation for the specified directory in UTC time.
'
' @param Path The directory to retrieve the creation time of.
' @return A CorDateTime object containing the creation time of the directory.
' @remarks The path can be relative or absolute.
'
Public Function GetCreationTimeUtc(ByRef Path As String) As CorDateTime
    Set GetCreationTimeUtc = File.GetCreationTimeUtc(Path)
End Function

''
' Returns the time the directory was last accessed in local time.
'
' @param Path The directory to retrieve the last time it was accessed.
' @return A CorDateTime object containing the time the directory was last accessed.
' @remarks The path can be relative or absolute.
'
Public Function GetLastAccessTime(ByRef Path As String) As CorDateTime
    Set GetLastAccessTime = File.GetLastAccessTime(Path)
End Function

''
' Returns the time the directory was last accessed in UTC time.
'
' @param Path The directory to retrieve the last time it was accessed.
' @return A CorDateTime object containing the time the directory was last accessed.
' @remarks The path can be relative or absolute.
'
Public Function GetLastAccessTimeUtc(ByRef Path As String) As CorDateTime
    Set GetLastAccessTimeUtc = File.GetLastAccessTimeUtc(Path)
End Function

''
' Returns the time the directory was last written to in local time.
'
' @param Path The directory to retrieve the last time it was written to.
' @return A CorDateTime object containing the last time the directory was written to.
' @remarks The path can relative or absolute.
'
Public Function GetLastWriteTime(ByRef Path As String) As CorDateTime
    Set GetLastWriteTime = File.GetLastWriteTime(Path)
End Function

''
' Returns the time the directory was last written to in UTC time.
'
' @param Path The directory to retrieve the last time it was written to.
' @return A CorDateTime object containing the last time the directory was written to.
' @remarks The path can relative or absolute.
'
Public Function GetLastWriteTimeUtc(ByRef Path As String) As CorDateTime
    Set GetLastWriteTimeUtc = File.GetLastWriteTimeUtc(Path)
End Function

''
' Returns the current directory the application is set to.
'
' @return The current set directory.
'
Public Function GetCurrentDirectory() As String
    Dim Result  As String
    Dim Size    As Long
    
    Result = String$(MAX_PATH, 0)
    Size = GetCurrentDirectoryW(Len(Result), Result)
    
    If Size = 0 Then
        ThrowIOError Err.LastDllError
    End If
    
    GetCurrentDirectory = Left$(Result, Size)
End Function

''
' Returns a list of all the directories found in the specified directory
' that matches the search pattern.
'
' @param Path The directory to find the requested directories.
' @param SearchPattern A pattern to compare all the directories against, returning those that match the pattern.
' @param SearchOption One of the enumeration values that specifies whether the search operation should include all subdirectories or only the current directory.
' @return An array of the matching directories. If no directories match the pattern, then
' an empty zero-length array.
' @remarks The path can be absolute or relative.
'
Public Function GetDirectories(ByRef Path As String, Optional ByRef SearchPattern As String = "*", Optional ByVal SearchOption As SearchOption = TopDirectoryOnly) As String()
    GetDirectories = GetFileSystemEntriesCore(Path, SearchPattern, DirectoryEntry, SearchOption)
End Function

''
' Returns the root directory of the specified path.
'
' @param Path The path to return the root of.
' @return The root of the specified directory.
' @remarks If the path is relative, then the current directory will
' be used as the root.
'
Public Function GetDirectoryRoot(ByRef Path As String) As String
    If Len(Path) = 0 Then _
        Error.Argument Argument_EmptyPath, "Path"
    
    If Statics.Path.IsPathRooted(Path) Then
        GetDirectoryRoot = Statics.Path.GetPathRoot(Path)
    Else
        Dim FullPath As String
        FullPath = Statics.Path.GetFullPath(Path)
        GetDirectoryRoot = Statics.Path.GetPathRoot(FullPath)
    End If
End Function

''
' Returns a list of all the files that match a pattern within the specified directory.
'
' @param Path The directory to search for the matching files.
' @param SearchPattern A pattern to match the files against.
' @param SearchOption One of the enumeration values that specifies whether the search operation should include all subdirectories or only the current directory.
' @return The list of matching file names. If no filenames matched, then an empty,
' zero-length array is returned.
' @remarks The path can be absolute or relative. The path is not included in the
' filename result.
'
Public Function GetFiles(ByRef Path As String, Optional ByRef SearchPattern As String = "*", Optional ByVal SearchOption As SearchOption = TopDirectoryOnly) As String()
    GetFiles = GetFileSystemEntriesCore(Path, SearchPattern, FileEntry, SearchOption)
End Function

''
' Returns both directories and files that match the specified pattern.
'
' @param Path The directory to search for the files and directories.
' @param SearchPattern The pattern the filenames and directories must match.
' @param SearchOption One of the enumeration values that specifies whether the search operation should include all subdirectories or only the current directory.
' @return The list of files and directories that matched the pattern. If no files or directories
' matched the pattern, then an empty, zero-length array is returned.
' @remarks The path can be absolute or relative.
'
Public Function GetFileSystemEntries(ByRef Path As String, Optional ByRef SearchPattern As String = "*", Optional SearchOption As SearchOption = TopDirectoryOnly) As String()
    GetFileSystemEntries = GetFileSystemEntriesCore(Path, SearchPattern, DirectoryEntry Or FileEntry, SearchOption)
End Function

''
' Returns a list of the logical drives installed on this machine.
'
' @return The list of logical drives on this machine.
'
Public Function GetLogicalDrives() As String()
    Dim Drives  As New ArrayList
    Dim Mask    As Long
    Dim Letter  As Long
    
    Letter = vbUpperAChar
    Mask = CorType.GetLogicalDrives
    
    Do While Mask > 0
        If Mask And 1 Then
            Drives.Add Chr$(Letter) & ":\"
        End If
        
        Letter = Letter + 1
        Mask = Mask \ 2
    Loop

    GetLogicalDrives = Drives.ToArray(vbString)
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
        Error.Argument Argument_EmptyPath, "Path"

    Dim FullPath    As String
    Dim DirName     As String
    
    FullPath = Statics.Path.GetFullPath(Path)
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
        Error.Argument Argument_EmptyFileName, "SourceDirName"
    If Len(DestDirName) = 0 Then _
        Error.Argument Argument_EmptyFileName, "DestDirName"
    
    Dim FullSourceDirName   As String
    Dim FullDestDirName     As String
    
    FullSourceDirName = Path.GetFullPath(SourceDirName)
    FullDestDirName = Path.GetFullPath(DestDirName)
    
    If CorString.Equals(FullSourceDirName, FullDestDirName, OrdinalIgnoreCase) Then _
        Error.IO IO_SourceDestMustBeDifferent
    If Path.GetPathRoot(FullSourceDirName) <> Path.GetPathRoot(FullDestDirName) Then _
        Error.IO IO_SourceDestMustHaveSameRoot
        
    If MoveFileW(FullSourceDirName, FullDestDirName) = BOOL_FALSE Then
        ThrowIOError Err.LastDllError, SourceDirName
    End If
End Sub

''
' Changes the current directory to the specified directory.
'
' @param Path The new directory to set as the current directory.
' @remarks The path can be absolute or relative.
'
Public Sub SetCurrentDirectory(ByRef Path As String)
    Dim FullPath As String
    
    If Len(Path) = 0 Then _
        Error.Argument Argument_EmptyPath, "Path"
        
    FullPath = Statics.Path.GetFullPath(Path)

    If SetCurrentDirectoryW(FullPath) = BOOL_FALSE Then
        ThrowIOError Err.LastDllError, FullPath
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
'   Shared
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ThrowIOError(ByVal ErrorCode As Long, Optional ByRef Path As String)
    If ErrorCode = ERROR_FILE_NOT_FOUND Then
        IOError.WinIOError ERROR_PATH_NOT_FOUND, Path
    Else
        IOError.WinIOError ErrorCode
    End If
End Sub

Private Function GetFileTimeUtc(ByRef Time As Variant) As Currency
    GetFileTimeUtc = Convert.ToDateTime(Time).ToFileTimeUtc
End Function

Private Function GetFileSystemEntriesCore(ByRef InitialPath As String, ByRef SearchPattern As String, ByVal EntriesToInclude As SystemEntryTypes, ByVal SearchOpt As SearchOption) As String()
    ValidateSearchPattern SearchPattern
    
    Dim Entry       As WIN32_FIND_DATAW
    Dim Entries     As New ArrayList
    Dim SubFolders  As New ArrayList
    Dim FullPath    As String
    
    FullPath = Path.GetFullPath(InitialPath)
    
    With SafeFindFirstFile(Path.Combine(FullPath, SearchPattern), Entry)
        If Not .IsInvalid Then
            Do
                Dim EntryName As String
                
                EntryName = GetFileNameFromFindData(Entry)
                
                If Len(EntryName) > 0 Then
                    Dim NewPath As String
                    
                    NewPath = Path.Combine(FullPath, EntryName)
                    
                    If PathIsDirectory(Entry) Then
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
            Loop While FindNextFileW(.DangerousGetHandle, Entry) <> BOOL_FALSE
        End If
        
        If IsIOError(Err.LastDllError) Then _
            IOError.WinIOError Err.LastDllError, InitialPath
    End With
    
    Dim SubFolder As Variant
    For Each SubFolder In SubFolders
        Entries.AddRange GetFileSystemEntriesCore(CStr(SubFolder), SearchPattern, EntriesToInclude, SearchOpt)
    Next

    GetFileSystemEntriesCore = Entries.ToArray(vbString)
End Function

Private Sub ValidateSearchPattern(ByRef Pattern As String)
    Dim Index As Long
    
    Index = 1
    
    Do
        Index = InStr(Index, Pattern, "..")
        
        Select Case Index
            Case 0
                Exit Do
            Case Len(Pattern) - 1
                Error.Argument Arg_InvalidSearchPattern
        End Select

        Select Case CorString.CharAt(Pattern, Index + 1)
            Case vbDirectorySeparatorChar, vbAltDirectorySeparatorChar
                Error.Argument Arg_InvalidSearchPattern
        End Select
        
        Index = Index + 2
    Loop
End Sub

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
            IOError.WinIOError Err.LastDllError, Path
        End If
    End With
End Sub

Private Function OpenHandle(ByRef Path As String) As SafeFileHandle
    Dim FullPath    As String
    Dim Root        As String
    Dim Handle      As SafeFileHandle
    
    FullPath = Statics.Path.GetFullPath(Path)
    Root = Statics.Path.GetPathRoot(FullPath)
    
    If (FullPath = Root) And (CorString.CharAt(Path, 1) = vbVolumeSeparatorChar) Then _
        Error.Argument Arg_PathIsAVolume
    
    Set Handle = Win32Native.SafeCreateFile(FullPath, WriteAccess, WriteShare, OpenExisting, FILE_FLAG_BACKUP_SEMANTICS)
    
    If Handle.IsInvalid Then _
        IOError.WinIOError Err.LastDllError, FullPath
        
    Set OpenHandle = Handle
End Function

Private Function GetFileNameFromFindData(ByRef Data As WIN32_FIND_DATAW) As String
    Dim FileName As String
    
    If Data.cFileName(0) > 0 Then
        FileName = SysAllocString(VarPtr(Data.cFileName(0)))
    Else
        FileName = SysAllocString(VarPtr(Data.cAlternateFileName(0)))
    End If
    
    If Not IsCurrentOrParentDirectory(FileName) Then
        GetFileNameFromFindData = FileName
    End If
End Function

Private Function IsCurrentOrParentDirectory(ByRef FileName As String) As Boolean
    IsCurrentOrParentDirectory = Len(FileName) <= 2 And CorString.AllChars(FileName, AddressOf IsPeriodCallback)
End Function

