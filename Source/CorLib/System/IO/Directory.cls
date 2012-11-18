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
'    Module: Directory
'

''
' Provides a set of static methods for manipulating and retrieving directory information.
'
Option Explicit
Private Const ALT_DIRECTORY_SEPARATOR_CHAR      As Integer = 47
Private Const DIRECTORY_SEPARATOR_CHAR          As Integer = 92
Private Const FILE_FLAG_BACKUP_SEMANTICS        As Long = &H2000000
Private Const ERROR_NO_MORE_FILES               As Long = 18&
Private Const DOUBLE_PERIOD                     As Long = &H2E002E

Private Enum SystemEntryTypes
    FileEntry = 1
    DirectoryEntry = 2
End Enum

' We want to use the variable name Path in the arguments like DotNet,
' so we will keep a reference to the one Path object for each access
' without having to type out the mStaticClasses portion.
Private mPath As Path



''
' Creates a directory and any subdirectories in the specified path.
'
' @param Path The directory path to be created.
' @return A DirectoryInfo object that represents the created directory.
' @remarks The path to be created can be a relative or absolute path.
'
Public Function CreateDirectory(ByVal Path As String) As DirectoryInfo
    Path = Trim$(Path)
    Call mPath.FixupPath(Path)
    
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "Path")
    
    Path = cString.TrimEnd(Path, mPath.DirectorySeparatorString)
    
    If Len(Path) > mPath.MaxPathLength Then _
        Throw New PathTooLongException
    If InStr(3, Path, mPath.VolumeSeparatorString) > 0 Then _
        Throw New NotSupportedException
    
    Dim DirectoryNames() As String
    DirectoryNames = Split(Path, mPath.DirectorySeparatorString)
    
    Dim i As Long
    i = cArray.GetLength(DirectoryNames)
    Do
        Dim TestDirectory As String
        TestDirectory = cString.Join(DirectoryNames, mPath.DirectorySeparatorString, 0, i)
        If Directory.Exists(TestDirectory) Then Exit Do
        i = i - 1
    Loop While i > 0
    
    Do While i <= UBound(DirectoryNames)
        i = i + 1
        Dim NewDirectory As String
        NewDirectory = cString.Join(DirectoryNames, mPath.DirectorySeparatorString, 0, i)
        If API.CreateDirectory(NewDirectory) = BOOL_FALSE Then IOError Err.LastDllError
    Loop
    
    Set CreateDirectory = Cor.NewDirectoryInfo(Path)
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
Public Sub Delete(ByVal Path As String, Optional ByVal Recursive As Boolean = False)
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath))
    
    Path = cString.TrimEnd(mPath.InternalGetFullPath(Path), "/\")
    
    If Recursive Then
        Dim FileHandle  As Long
        Dim Data        As WIN32_FIND_DATA
        
        FileHandle = API.FindFirstFile(Path & "\*", Data)
        If FileHandle <> INVALID_HANDLE Then
            ' Simply loop through any files or directories
            ' within this directory and delete them.
            Do
                Dim FileName As String
                FileName = GetFileNameFromFindData(Data)
                
                If Len(FileName) > 0 Then
                    Dim FullPath As String
                    FullPath = mPath.Combine(Path, FileName)
                    
                    If Data.dwFileAttributes And FileAttributes.DirectoryAttr Then
                        Call Delete(FullPath, True)
                    Else
                        If API.DeleteFile(FullPath) = BOOL_FALSE Then IOError Err.LastDllError, FullPath
                    End If
                End If
            Loop While API.FindNextFile(FileHandle, Data)
            
            Dim Result As Long
            Result = Err.LastDllError    ' get this before FindClose can change it.
            
            Call FindClose(FileHandle)
            If Result <> ERROR_NO_MORE_FILES Then IOError Result, FullPath
        End If
    End If
    
    If API.RemoveDirectory(Path) = BOOL_FALSE Then
        If Err.LastDllError <> ERROR_PATH_NOT_FOUND Then IOError Err.LastDllError, Path
    End If
End Sub

''
' Returns the time of creation for the specified directory in local time.
'
' @param Path The directory to retrieve the creation time of.
' @return A cDateTime object containing the creation time of the directory.
' @remarks The path can be relative or absolute.
'
Public Function GetCreationTime(ByVal Path As String) As cDateTime
    Set GetCreationTime = File.GetCreationTime(Path)
End Function

''
' Returns the time of creation for the specified directory in UTC time.
'
' @param Path The directory to retrieve the creation time of.
' @return A cDateTime object containing the creation time of the directory.
' @remarks The path can be relative or absolute.
'
Public Function GetCreationTimeUtc(ByVal Path As String) As cDateTime
    Set GetCreationTimeUtc = File.GetCreationTimeUtc(Path)
End Function

''
' Returns the time the directory was last accessed in local time.
'
' @param Path The directory to retrieve the last time it was accessed.
' @return A cDateTime object containing the time the directory was last accessed.
' @remarks The path can be relative or absolute.
'
Public Function GetLastAccessTime(ByVal Path As String) As cDateTime
    Set GetLastAccessTime = File.GetLastAccessTime(Path)
End Function

''
' Returns the time the directory was last accessed in UTC time.
'
' @param Path The directory to retrieve the last time it was accessed.
' @return A cDateTime object containing the time the directory was last accessed.
' @remarks The path can be relative or absolute.
'
Public Function GetLastAccessTimeUtc(ByVal Path As Stream) As cDateTime
    Set GetLastAccessTimeUtc = File.GetLastAccessTimeUtc(Path)
End Function

''
' Returns the time the directory was last written to in local time.
'
' @param Path The directory to retrieve the last time it was written to.
' @return A cDateTime object containing the last time the directory was written to.
' @remarks The path can relative or absolute.
'
Public Function GetLastWriteTime(ByVal Path As String) As cDateTime
    Set GetLastWriteTime = File.GetLastWriteTime(Path)
End Function

''
' Returns the time the directory was last written to in UTC time.
'
' @param Path The directory to retrieve the last time it was written to.
' @return A cDateTime object containing the last time the directory was written to.
' @remarks The path can relative or absolute.
'
Public Function GetLastWriteTimeUtc(ByVal Path As String) As cDateTime
    Set GetLastWriteTimeUtc = File.GetLastWriteTimeUtc(Path)
End Function

''
' Returns the current directory the application is set to.
'
' @return The current set directory.
'
Public Function GetCurrentDirectory() As String
    Dim Ret As String
    Ret = String$(Path.MaxPathLength, 0)
    
    Dim size As Long
    size = API.GetCurrentDirectory(Len(Ret), Ret)
    If size = 0 Then IOError Err.LastDllError
    
    GetCurrentDirectory = Left$(Ret, size)
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
Public Function GetDirectories(ByVal Path As String, Optional ByVal SearchPattern As String = "*") As String()
    GetDirectories = InternalGetFiles(Path, SearchPattern, DirectoryEntry)
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
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "Path")
    
    If Not mPath.IsPathRooted(Path) Then
        Path = mPath.GetFullPath(Path)
    End If
    
    GetDirectoryRoot = mPath.GetPathRoot(Path)
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
Public Function GetFiles(ByVal Path As String, Optional ByVal SearchPattern As String = "*") As String()
    GetFiles = InternalGetFiles(Path, SearchPattern, FileEntry)
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
Public Function GetFileSystemEntries(ByVal Path As String, Optional ByVal SearchPattern As String = "*") As String()
    GetFileSystemEntries = InternalGetFiles(Path, SearchPattern, DirectoryEntry Or FileEntry)
End Function

''
' Returns a list of the logical drives installed on this machine.
'
' @return The list of logical drives on this machine.
'
Public Function GetLogicalDrives() As String()
    Dim Ret() As String
    ReDim Ret(25) As String ' just allocate the maximum letters possible.
    
    Dim mask As Long
    mask = VBCorType.GetLogicalDrives
    
    Dim letter As Long
    letter = vbUpperA
    
    Dim i As Long
    Do While mask <> 0
        If mask And 1 Then
            Ret(i) = Chr$(letter) & ":\"
            i = i + 1
        End If
        letter = letter + 1
        mask = mask \ 2
    Loop
    
    ReDim Preserve Ret(0 To i - 1)
    GetLogicalDrives = Ret
End Function

''
' Returns the parent of the specified directory.
'
' @param Path The directory to retrieve the parent of.
' @return A DirectoryInfo object representing the parent directory.
' @remarks The path can be absolute or relative.
'
Public Function GetParent(ByVal Path As String) As DirectoryInfo
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "Path")

    Call mPath.VerifyPath(Path)
    Path = mPath.GetFullPath(Path)
    
    Dim DirName As String
    DirName = mPath.GetDirectoryName(Path)
    If cString.IsNull(DirName) Then Exit Function
    Set GetParent = Cor.NewDirectoryInfo(DirName)
End Function

''
' Moves a directory and all of its contents to another location.
'
' @param SourceDirectory The directory to be moved.
' @param DestinationDirectory The directory to be moved to.
' @remarks The destination directory is the source directory renamed
' after the moving of the source directory.
'
Public Sub Move(ByVal SourceDirectory As String, ByVal DestinationDirectory As String)
    If Len(SourceDirectory) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "SourceDirectory")
    If Len(DestinationDirectory) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "DestinationDirectory")
    
    SourceDirectory = LCase$(Path.InternalGetFullRootedPath(SourceDirectory))
    DestinationDirectory = LCase$(Path.InternalGetFullRootedPath(DestinationDirectory))
    
    If SourceDirectory = DestinationDirectory Then _
        Throw Cor.NewIOException("Source and Destination directories must be different.")
    If Path.GetPathRoot(SourceDirectory) <> Path.GetPathRoot(DestinationDirectory) Then _
        Throw Cor.NewIOException("Source and Destination directories must be on the same root drive.")
        
    If API.MoveFile(SourceDirectory, DestinationDirectory) = BOOL_FALSE Then
        Dim Result  As Long
        Result = Err.LastDllError
        If Result = ERROR_FILE_NOT_FOUND Then
            IOError ERROR_PATH_NOT_FOUND, SourceDirectory
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
Public Sub SetCurrentDirectory(ByVal Path As String)
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "Path")
    
    Call mPath.VerifyPath(Path)
    If Not Exists(Path) Then _
        Throw New DirectoryNotFoundException
    
    If API.SetCurrentDirectory(Path) = BOOL_FALSE Then IOError Err.LastDllError, Path
End Sub

''
' Sets the creation time of the specified directory to the local time.
'
' @param Path The directory to set the creation time for.
' @param CreationTime A Date or cDateTime object of the new time.
' @remarks The time will have the UTC offset removed before setting the
' the directory to it. When the time is read, it will be in local time to the timezone.
'
Public Sub SetCreationTime(ByVal Path As String, ByVal CreationTime As Variant)
    Call InternalSetFileTime(Path, cDateTime.GetcDateTime(CreationTime).ToFileTime)
End Sub

''
' Sets the creation time of the specified directory to UTC time.
'
' @param Path The directory to set the creation time for in UTC.
' @param CreationTime A Date or cDateTime object of the new UTC time.
' @remarks The directory is set to the actual time passed in.
'
Public Sub SetCreationTimeUtc(ByVal Path As String, ByVal CreationTime As Variant)
    Call InternalSetFileTime(Path, cDateTime.GetcDateTime(CreationTime).ToFileTimeUtc)
End Sub

''
' Sets the time the directory was last accessed in local time.
'
' @param Path The directory to set the new time for.
' @param LastAccessTime A Date or cDateTime object of the new time.
' @remarks The UTC offset is removed from the time before setting the directory.
'
Public Sub SetLastAccessTime(ByVal Path As String, ByVal LastAccessTime As Variant)
    Call InternalSetFileTime(Path, , cDateTime.GetcDateTime(LastAccessTime).ToFileTime)
End Sub

''
' Sets the time the directory was last accessed in UTC time.
'
' @param Path The directory to set the new time for.
' @param LastAccessTime A Date or cDateTime object of the new time.
' @remarks The directory is set to the actual time passed in.
'
Public Sub SetLastAccessTimeUtc(ByVal Path As String, ByVal LastAccessTime As Variant)
    Call InternalSetFileTime(Path, , cDateTime.GetcDateTime(LastAccessTime).ToFileTimeUtc)
End Sub

''
' Sets the time the directory was last written to in local time.
'
' @param Path The directory to set the new time for.
' @param LastWriteTime A Date or cDateTime object of the new time.
'
Public Sub SetLastWriteTime(ByVal Path As String, ByVal LastWriteTime As Variant)
    Call InternalSetFileTime(Path, , , cDateTime.GetcDateTime(LastWriteTime).ToFileTime)
End Sub

''
' Sets the time the directory was last written to in UTC time.
'
' @param Path The directory to set the new time for.
' @param LastWriteTime A Date or cDateTime object of the new time.
'
Public Sub SetLastWriteTimeUtc(ByVal Path As String, ByVal LastWriteTime As Variant)
    Call InternalSetFileTime(Path, , , cDateTime.GetcDateTime(LastWriteTime).ToFileTimeUtc)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifySearchPattern(ByRef s As String)
    Dim i As Long
    i = InStr(s, "..")
    If i > 0 Then
        If i = Len(s) - 1 Then Throw Cor.NewArgumentException("Cannot end search pattern with '..'")
        Select Case Asc(Mid$(s, i + 2, 1))
            Case DIRECTORY_SEPARATOR_CHAR, ALT_DIRECTORY_SEPARATOR_CHAR
                Throw Cor.NewArgumentException("'..' cannot be followed immediately by a directory separator.")
        End Select
    End If
End Sub

Private Function InternalGetFiles(ByRef sPath As String, ByRef SearchPattern As String, ByVal EntryTypes As SystemEntryTypes) As String()
    Dim FullPath As String
    FullPath = Path.InternalGetFullPath(sPath)
    
    Call VerifySearchPattern(SearchPattern)
    
    Dim Data        As WIN32_FIND_DATA
    Dim FileHandle  As Long
    FileHandle = API.FindFirstFile(Path.Combine(sPath, SearchPattern), Data)
    
    Dim list As New ArrayList
    
    Dim include As Boolean
    If FileHandle <> INVALID_HANDLE Then
        Do
            If Data.dwFileAttributes And FileAttributes.DirectoryAttr Then
                include = (EntryTypes And SystemEntryTypes.DirectoryEntry)
            Else
                include = (EntryTypes And SystemEntryTypes.FileEntry)
            End If
            
            If include Then
                Dim s As String
                s = GetFileNameFromFindData(Data)
                If Len(s) > 0 Then Call list.Add(Path.Combine(sPath, s))
            End If
        Loop While API.FindNextFile(FileHandle, Data) <> BOOL_FALSE
    End If
    
    Dim e As Long
    e = Err.LastDllError    ' cache the error before FindClose changes it.
    
    Call FindClose(FileHandle)
    
    Select Case e
        Case 0, ERROR_NO_MORE_FILES, ERROR_FILE_NOT_FOUND   ' ignore these errors.
        Case Else: IOError e, sPath
    End Select

    InternalGetFiles = list.ToArray(vbString)
End Function

Private Sub InternalSetFileTime(ByRef sPath As String, Optional ByVal CreationTime As Variant, Optional ByVal LastAccessTime As Variant, Optional ByVal LastWriteTime As Variant)
    Dim cnt As Long
    Dim lat As Long
    Dim lwt As Long
    
    If Not IsMissing(CreationTime) Then cnt = VarPtr(CreationTime) + VARIANTDATA_OFFSET
    If Not IsMissing(LastAccessTime) Then lat = VarPtr(LastAccessTime) + VARIANTDATA_OFFSET
    If Not IsMissing(LastWriteTime) Then lwt = VarPtr(LastWriteTime) + VARIANTDATA_OFFSET
    
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    Call Path.VerifyPath(sPath)
    
    Dim DirHandle As Long
    DirHandle = API.CreateFile(sPath, FileAccess.WriteAccess, FileShare.None, 0, FileMode.OpenExisting, FILE_FLAG_BACKUP_SEMANTICS, 0)
    If DirHandle = INVALID_HANDLE Then IOError Err.LastDllError, sPath
    If SetFileTime(DirHandle, ByVal cnt, ByVal lat, ByVal lwt) = BOOL_FALSE Then
        Close CloseHandle(DirHandle)
        IOError Err.LastDllError, sPath
    End If
    Call CloseHandle(DirHandle)
End Sub

Private Function GetFileNameFromFindData(ByRef Data As WIN32_FIND_DATA) As String
    Dim Ret As String
    Ret = Data.cFileName
    If Len(Ret) = 0 Then Ret = Data.cAlternateFileName
    
    If (Data.dwFileAttributes And FileAttributes.DirectoryAttr) = FileAttributes.DirectoryAttr Then
        Select Case Len(Ret)
            Case 1: If Asc(Ret) = vbPeriod Then Exit Function
            Case 2: If AsLong(ByVal StrPtr(Ret)) = DOUBLE_PERIOD Then Exit Function
        End Select
    End If
    
    GetFileNameFromFindData = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mPath = modStaticClasses.Path
End Sub
