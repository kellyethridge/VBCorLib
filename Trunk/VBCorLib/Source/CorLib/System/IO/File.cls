VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File"
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
'    Module: File
'

''
' Provides methods for manipulating and retrieving file information.
'
' @remarks This class cannot be directly created. To access the methods
' use the variable name directly.
'
' <pre>
' Debug.Print File.Exists("MyFile.txt")
' </pre>
'
Option Explicit

Private mDirectorySeparators() As Integer


''
' Opens a file to be written to starting at the end of the file.
'
' @param Path The name of the file to be appended to.
' @return A StreamWriter object used to write to the file.
' @remarks The StreamWriter uses a UTF8Encoding declared as Cor.NewUTF8Encoding(False, True).
' @see StreamWriter
'
Public Function AppendText(ByRef Path As String) As StreamWriter
    Set AppendText = Cor.NewStreamWriter(Path, True)
End Function

''
' Append a text string to the end of a specified file using the supplied encoding.
'
' @param Path The file to append the text to.
' @param Contents The string to append to the file.
' @param Encoding The encoding used to encode the contents (Default UTF8).
' @remarks This method opens a file, appends the contents then closes the
' file in the call. If multiple lines need to be appended, then using this
' method may have performance problems.
' <p>If the file doesn't exist it will be created. If the file already
' exists, it will be overwritten.</p>
'
Public Sub AppendAllText(ByRef Path As String, ByRef Contents As String, Optional ByVal Encoding As Encoding)
    With Cor.NewStreamWriter(Path, True, Encoding)
        .WriteValue Contents
    End With
End Sub

''
' Copies a file.
'
' @param SourceFileName The file to be copied.
' @param DestinationFileName The location and name of the copied file.
' @param OverWrite Indicates if the file already exists it should be overwritten.
'
Public Sub Copy(ByVal SourceFileName As String, ByVal DestinationFileName As String, Optional ByVal OverWrite As Boolean = False)
    If Len(SourceFileName) = 0 Then _
        Error.Argument Argument_EmptyFileName, "SourceFileName"
    If Len(DestinationFileName) = 0 Then _
        Error.Argument Argument_EmptyFileName, "DestinationFileName"
    If Path.HasIllegalCharacters(SourceFileName) Then _
        Error.Argument Argument_InvalidCharsInPath, "SourceFileName"
    If Path.HasIllegalCharacters(DestinationFileName) Then _
        Error.Argument Argument_InvalidCharsInPath, "DestinationFileName"
        
    SourceFileName = Path.GetFullPath(cString.TrimEnd(SourceFileName, mDirectorySeparators))
    DestinationFileName = Path.GetFullPath(cString.TrimEnd(DestinationFileName, mDirectorySeparators))
    
    If cString.EqualsIgnoreCase(SourceFileName, DestinationFileName) Then
        Exit Sub
    End If
    
    If Api.CopyFile(SourceFileName, DestinationFileName, Not OverWrite) = BOOL_FALSE Then
        Dim DllError As Long
        DllError = Err.LastDllError
        If DllError <> ERROR_FILE_EXISTS Then
            If Not File.Exists(SourceFileName) Then _
                IOError ERROR_FILE_NOT_FOUND, SourceFileName
            If Directory.Exists(DestinationFileName) Then _
                Throw Cor.NewIOException(Environment.GetResourceString(ResourceString.IOException_DirectoryExists))
        End If
        IOError DllError, SourceFileName
    End If
End Sub

''
' Creates a file. If the file already exists, it is overwritten.
'
' @param Path The new file to be created.
' @param BufferSize The size of the internal buffer used byte the FileStream object.
' @return A FileStream object to be used for reading and writing to the file.
' @remarks The FileStream is declared as Cor.NewFileStream(Path, FileMode.Create, FileAccess.ReadWriteAccess, FileShare.None, BufferSize).
' @see FileStream
'
Public Function Create(ByVal Path As String, Optional ByVal BufferSize As Long = 4096) As FileStream
    Set Create = Cor.NewFileStream(Path, FileMode.Create, FileAccess.ReadWriteAccess, FileShare.None, BufferSize)
End Function

''
' Creates a file using a UTF8Encoding. If the file already exists, then it is overwritten.
'
' @param Path The new file to be created.
' @return A StreamWriter to be used to write to the new file.
' @remarks The StreamWriter is declared as Cor.NewStreamWriter(Path).
' @see StreamWriter
'
Public Function CreateText(ByVal Path As String) As StreamWriter
    Set CreateText = Cor.NewStreamWriter(Path)
End Function

''
' Deletes a file.
'
' @param Path The file to be deleted.
' @remarks If the file did not exist, nothing happens.
'
Public Sub Delete(ByVal Path As String)
    If Len(Path) = 0 Then _
        Error.Argument Argument_EmptyPath, "Path"
    
    VBCorLib.Path.VerifyPath Path, True
    
    If Api.DeleteFile(Path) = BOOL_FALSE Then
        Dim Result As Long
        Result = Err.LastDllError
        If Result <> ERROR_FILE_NOT_FOUND Then
            IOError Result, Path
        End If
    End If
End Sub

''
' Determines if a file exists.
'
' @param Path The file to check for existence.
' @return Whether the file exists or not.
'
Public Function Exists(ByVal Path As String) As Boolean
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    
    If GetFileData(Path, Data) = NO_ERROR Then
        Exists = Not CBool(Data.dwFileAttributes And FileAttributes.DirectoryAttr)
    End If
End Function

''
' Returns the file attributes (ReadOnly, Hidden, ...)
'
' @param Path The file to retrieve the attributes for.
' @return A value with individual bits representing if an attribute is applied to the file.
' @remarks To determine if a specific attribute is being used, the return value will
' need to be ANDed with the specific FileAttributes flag.<br>
' <pre>
' If Attr And FileAttributes.ReadOnly Then
'     ''... file is readonly
' End If
' </pre>
'
Public Function GetAttributes(ByRef Path As String) As FileAttributes
    Dim FullPath As String
    FullPath = Statics.Path.GetFullPath(Path)
    
    GetAttributes = Api.GetFileAttributes(FullPath)
    If GetAttributes = INVALID_FILE_ATTRIBUTES Then
        IOError Err.LastDllError, FullPath
    End If
End Function

''
' Returns the time the file was created in local time.
'
' @param Path The file to retrieve the creation time of.
' @return A cDateTime object representing the creation time.
'
Public Function GetCreationTime(ByRef Path As String) As cDateTime
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    FillFileData Path, Data
    Set GetCreationTime = cDateTime.FromFileTime(Data.ftCreationTime)
End Function

''
' Returns the time the file was created in UTC time.
'
' @param Path The file to retrieve the creation time of.
' @return A cDateTime object representing the creation time.
'
Public Function GetCreationTimeUtc(ByRef Path As String) As cDateTime
    Set GetCreationTimeUtc = GetCreationTime(Path).ToUniversalTime
End Function

''
' Returns the time the file was accessed in local time.
'
' @param Path The file to retrieve the last access time for.
' @return A cDateTime object that represents the last time the file was accessed.
'
Public Function GetLastAccessTime(ByRef Path As String) As cDateTime
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    FillFileData Path, Data
    Set GetLastAccessTime = cDateTime.FromFileTime(Data.ftLastAccessTime)
End Function

''
' Returns The last time a file was accessed in UTC time.
'
' @param Path The file to get the last accessed time for.
' @return A cDateTime object representing the last time the file was accessed.
'
Public Function GetLastAccessTimeUtc(ByRef Path As String) As cDateTime
    Set GetLastAccessTimeUtc = GetLastAccessTime(Path).ToUniversalTime
End Function

''
' Returns the last time the file was written to in local time.
'
' @param Path The file to retrieve the last written time for.
' @return A cDateTime object representing the last time the file was written to.
'
Public Function GetLastWriteTime(ByRef Path As String) As cDateTime
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    FillFileData Path, Data
    Set GetLastWriteTime = cDateTime.FromFileTime(Data.ftLastWriteTime)
End Function

''
' Returns the last time the file was written to in UTC time.
'
' @param sPath The file to retrieve the last written time for.
' @return A cDateTime object representing the last time a file was written to.
'
Public Function GetLastWriteTimeUtc(ByVal sPath As String) As cDateTime
    Set GetLastWriteTimeUtc = GetLastWriteTime(sPath).ToUniversalTime
End Function

''
' Moves a file from one location to another.
'
' @param SourceFileName The file to be moved.
' @param DestinationFileName The location and filename the file is to be moved to.
'
Public Sub Move(ByRef SourceFileName As String, ByVal DestinationFileName As String)
    If Len(SourceFileName) = 0 Then _
        Error.Argument Argument_EmptyPath, "SourceFileName"
    If Len(DestinationFileName) = 0 Then _
        Error.Argument Argument_EmptyPath, "DestinationFileName"

    Call Path.VerifyPath(SourceFileName)
    Call Path.VerifyPath(DestinationFileName)
    
    If Not File.Exists(SourceFileName) Then _
        IOError ERROR_FILE_NOT_FOUND, SourceFileName
    
    If Api.MoveFile(SourceFileName, DestinationFileName) = BOOL_FALSE Then IOError Err.LastDllError, DestinationFileName
End Sub

''
' Opens a file with a FileStream
'
' @param Path The file to be opened.
' @param Mode The mode in which to open the file.
' @param Access The type of access allowed by the FileStream.
' @param Share The access allowed for other processes.
' @return A FileStream object of the specified file.
'
Public Function OpenFile(ByRef Path As String, ByVal Mode As FileMode, Optional ByVal Access As FileAccess = -1, Optional ByVal Share As FileShare = FileShare.None) As FileStream
    Set OpenFile = Cor.NewFileStream(Path, Mode, Access, Share)
End Function

''
' Opens a file for reading as a FileStream.
'
' @param Path The file to be opened in read mode.
' @return A FileStream used to read from the file.
' @remarks The FileStream object is declared as Cor.NewFileStream(Path, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadShare).
'
Public Function OpenRead(ByRef Path As String) As FileStream
    Set OpenRead = Cor.NewFileStream(Path, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadShare)
End Function

''
' Opens a file for reading through a StreamReader.
'
' @param Path The file to be opened for reading.
' @return A StreamReader used to read from the file with the default encoding (UTF8Encoding).
' @remarks The StreamReader is declared as Cor.NewStreamReader(Path).
'
Public Function OpenText(ByRef Path As String) As StreamReader
    Set OpenText = Cor.NewStreamReader(Path)
End Function

''
' Opens a file for writing as a FileStream.
'
' @param Path The file to be opend for writing to.
' @return A FileStream object used to write to the specified file.
' @remarks The FileStream object is declared as Cor.NewFileStream(Path, FileMode.OpenOrCreate, FileAccess.WriteAccess, FileShare.None).
'
Public Function OpenWrite(ByRef Path As String) As FileStream
    Set OpenWrite = Cor.NewFileStream(Path, FileMode.OpenOrCreate, FileAccess.WriteAccess, FileShare.None)
End Function

''
' Sets the file attributes (ReadOnly, Hidden, ...)
'
' @param Path The file to set the attributes of.
' @param Attributes The new attributes to set the file to.
' @remarks In order to set a single attribute without unsetting an existing
' attributes, they will first need to be read from the file (GetAttributes), then
' bit manipulated to create the resulting set of desired attributes before setting.
'
Public Sub SetAttributes(ByRef Path As String, ByVal Attributes As FileAttributes)
    If Len(Path) = 0 Then _
        Error.Argument Argument_EmptyPath, "Path"
    
    Statics.Path.CheckValidPathAndFixSlashes Path
    
    If Api.SetFileAttributes(Path, Attributes) = BOOL_FALSE Then
        IOError Err.LastDllError, Path
    End If
End Sub

''
' Sets the time of creation for a file in local time.
'
' @param Path The file to set the time of creation for.
' @param CreationTime A Date or cDateTime object of the new time.
'
Public Sub SetCreationTime(ByVal Path As String, ByVal CreationTime As Variant)
    InternalSetFileTime Path, cDateTime.GetcDateTime(CreationTime).ToFileTime
End Sub

''
' Sets the time of creation for a file in UTC time.
'
' @param Path The file to set the time of creation for.
' @param CreationTime A Date or cDateTime object of the new time.
'
Public Sub SetCreationTimeUtc(ByVal Path As String, ByVal CreationTime As Variant)
    InternalSetFileTime Path, cDateTime.GetcDateTime(CreationTime).ToFileTimeUtc
End Sub

''
' Sets the last time the file was accessed in local time.
'
' @param Path The file to set the last access time for.
' @param LastAccessTime A Date or cDateTime object of the new time.
'
Public Sub SetLastAccessTime(ByVal Path As String, ByVal LastAccessTime As Variant)
    InternalSetFileTime Path, , cDateTime.GetcDateTime(LastAccessTime).ToFileTime
End Sub

''
' Sets the last time the file was accessed in UTC time.
'
' @param Path The file to set the last access time for.
' @param LastAccessTime A Date or cDateTime object of the new time.
'
Public Sub SetLastAccessTimeUtc(ByVal Path As String, ByVal LastAccessTime As Variant)
    InternalSetFileTime Path, , cDateTime.GetcDateTime(LastAccessTime).ToFileTimeUtc
End Sub

''
' Sets the last time the file was written to in local time.
'
' @param sPath The file to set the last written time for.
' @param LastWriteTime A Date or cDateTime object of the new time.
'
Public Sub SetLastWriteTime(ByVal Path As String, ByVal LastWriteTime As Variant)
    InternalSetFileTime Path, , , cDateTime.GetcDateTime(LastWriteTime).ToFileTime
End Sub

''
' Sets the last time the file was written to in UTC time.
'
' @param sPath The file to set the last written time for.
' @param LastWriteTime A Date or cDateTime object of the new time.
'
Public Sub SetLastWriteTimeUtc(ByVal Path As String, ByVal LastWriteTime As Variant)
    InternalSetFileTime Path, , , cDateTime.GetcDateTime(LastWriteTime).ToFileTimeUtc
End Sub

''
' Reads the entire contents of a file and returns it as a String.
'
' @param Path The path and filename to read in.
' @param Encoding The encoding to be used in reading the file.
' @return A string containing the contents of the file.
'
Public Function ReadAllText(ByVal Path As String, Optional ByVal Encoding As Encoding) As String
    With Cor.NewStreamReader(Path, Encoding)
        ReadAllText = .ReadToEnd
    End With
End Function

''
' Reads an entire file into a byte array.
'
' @param Path The path and filename to read into the byte array.
' @return A byte array containing the contents of the file.
'
Public Function ReadAllBytes(ByRef Path As String) As Byte()
    Const MaxLength As Long = &H7FFFFFFF
    Dim Bytes() As Byte
    Dim Length  As Long
        
    With Cor.NewFileStream(Path, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadWriteShare)
        Length = .Length
        If Length > MaxLength Then _
            Throw Cor.NewIOException(Environment.GetResourceString(ResourceString.IOException_FileTooLong2GB))
            
        If Length > 0 Then
            ReDim Bytes(0 To Length - 1)
            .ReadBlock Bytes, 0, Length
        Else
            Bytes = Cor.NewBytes()
        End If
    End With
    
    ReadAllBytes = Bytes
End Function

''
' Opens a file, reads all lines in a file into an array and closes the files.
'
' @param Path The path and filename of the file to read in.
' @param Encoding The encoding to use to decode the file into text characters.
' @return A String array containing all of the lines in the file.
' @remarks The end of a line is indicated when either a Return (13), LineFeed (10),
' or a Return-LineFeed combination have been encountered.
'
Public Function ReadAllLines(ByVal Path As String, Optional ByVal Encoding As Encoding) As String()
    Dim Lines() As String
    Dim Index   As Long
    
    If Len(Path) = 0 Then _
        Error.Argument Argument_EmptyPath, "Path"
    
    ReDim Lines(0 To 15)
    With Cor.NewStreamReader(Path, Encoding)
        Do
            If Index > UBound(Lines) Then
                ReDim Preserve Lines(0 To Index * 2)
            End If
            
            Lines(Index) = .ReadLine
            If StrPtr(Lines(Index)) = vbNullPtr Then
                Exit Do
            End If
            
            Index = Index + 1
        Loop
    End With
    
    If Index > 0 Then
        ReDim Preserve Lines(0 To Index - 1)
    Else
        Lines = Cor.NewStrings()
    End If
    
    ReadAllLines = Lines
End Function

''
' Opens a files, writes out all contents to the file, then closes the file.
'
' @param Path The path and filename of the file to write to.
' @param Contents The contents to write to the file.
' @param Encoding The encoding to be used when writing to the file.
' @Remarks If the file does not exist, it will be created. If the file already
' exists, it will be overwritten.
'
Public Sub WriteAllText(ByRef Path As String, ByVal Contents As String, Optional ByVal Encoding As Encoding)
    With Cor.NewStreamWriter(Path, Encoding:=Encoding)
        .WriteValue Contents
    End With
End Sub

''
' Opens a file, writes all bytes to the file, then closes the file.
'
' @param Path The path and filename of the file to write to.
' @param Bytes The bytes to be written to the file.
' @remarks If the file doesn't exist it will be created. If the file already
' exists, it will be overwritten.
'
Public Sub WriteAllbytes(ByRef Path As String, ByRef Bytes() As Byte)
    If CorArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException("Bytes", Environment.GetResourceString(ArgumentNull_Array))
    
    With Cor.NewFileStream(Path, FileMode.Create)
        .WriteBlock Bytes, LBound(Bytes), CorArray.GetLength(Bytes)
    End With
End Sub

''
' Opens a file, writes all strings to the file with appended new line values, then closes the file.
'
' @param Path The path and filename of the file to write to.
' @param Contents The strings to be written to the file.
' @param Encoding The encoding to be used when writing to the file.
' @remarks If the file doesn't exist it will be created. If the file already
' exists, it will be overwritten.
'
Public Sub WriteAllLines(ByVal Path As String, ByRef Contents As Variant, Optional ByVal Encoding As Encoding)
    With Cor.NewStreamWriter(Path, Encoding:=Encoding)
        Dim Line As Variant
        
        For Each Line In Contents
            .WriteLine Line
        Next
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetFileData(ByVal p As String, ByRef Data As WIN32_FILE_ATTRIBUTE_DATA) As Long
    p = cString.TrimEnd(p, mDirectorySeparators)
    
    Call Path.VerifyPath(p, True)
    If Api.GetFileAttributesEx(p, 0, Data) = BOOL_FALSE Then
        GetFileData = GetFileDataFallback(p, Data)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetFileDataFallback(ByRef Path As String, ByRef Data As WIN32_FILE_ATTRIBUTE_DATA) As Long
    Dim FileData As WIN32_FIND_DATA
    
    With SafeFindFirstFile(Path, FileData)
        If Not .IsInvalid Then
            Data.dwFileAttributes = FileData.dwFileAttributes
            Data.ftCreationTime = FileData.ftCreationTime
            Data.ftLastAccessTime = FileData.ftLastAccessTime
            Data.ftLastWriteTime = FileData.ftLastWriteTime
            Data.nFileSizeHigh = FileData.nFileSizeHigh
            Data.nFileSizeLow = FileData.nFileSizeLow
        Else
            GetFileDataFallback = Err.LastDllError
        End If
    End With
End Function

Private Sub InternalSetFileTime(ByRef Path As String, Optional ByVal CreationTime As Variant, Optional ByVal LastAccessTime As Variant, Optional ByVal LastWriteTime As Variant)
    Dim CreationTimePtr As Long
    Dim LastAccessTimePtr As Long
    Dim LastWriteTimePtr As Long
    
    If Not IsMissing(CreationTime) Then
        CreationTimePtr = VarPtr(CreationTime) + 8
    End If
    
    If Not IsMissing(LastAccessTime) Then
        LastAccessTimePtr = VarPtr(LastAccessTime) + 8
    End If
    
    If Not IsMissing(LastWriteTime) Then
        LastWriteTimePtr = VarPtr(LastWriteTime) + 8
    End If
    
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "Path")
        
    VBCorLib.Path.CheckValidPath Path
    
    Dim FileHandle As Long
    FileHandle = Api.CreateFile(Path, FileAccess.WriteAccess, FileShare.None, 0, FileMode.OpenExisting, FILE_ATTRIBUTE_NORMAL, 0)
    If FileHandle = INVALID_HANDLE Then _
        IOError Err.LastDllError, Path
        
    If SetFileTime(FileHandle, ByVal CreationTimePtr, ByVal LastAccessTimePtr, ByVal LastWriteTimePtr) = BOOL_FALSE Then
        Dim e As Long
        e = Err.LastDllError
        CloseHandle FileHandle
        IOError e, Path
    End If
    
    CloseHandle FileHandle
End Sub

Private Sub FillFileData(ByRef Path As String, ByRef Data As WIN32_FILE_ATTRIBUTE_DATA)
    If Len(Path) = 0 Then _
        Error.Argument Argument_EmptyPath, "Path"
    
    Dim DllError As Long
    DllError = GetFileData(Path, Data)
    If DllError <> NO_ERROR Then
        IOError DllError, Path
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mDirectorySeparators = Cor.NewIntegers(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
End Sub
