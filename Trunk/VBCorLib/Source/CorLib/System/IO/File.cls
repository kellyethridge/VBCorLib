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
Public Function AppendText(ByVal Path As String) As StreamWriter
    Set AppendText = Cor.NewStreamWriter(Path, Append:=True)
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
Public Sub AppendAllText(ByVal Path As String, ByVal Contents As String, Optional ByVal Encoding As Encoding)
    Dim sw As StreamWriter
    
    Set sw = Cor.NewStreamWriter(Path, Encoding, Append:=True)
    Call sw.WriteValue(Contents)
    Call sw.CloseWriter
End Sub

''
' Copies a file.
'
' @param SourceFileName The file to be copied.
' @param DestinationFileName The location and name of the copied file.
' @param OverWrite Indicates if the file already exists it should be overwritten.
'
Public Sub Copy(ByVal SourceFileName As String, ByVal DestinationFileName As String, Optional ByVal OverWrite As Boolean = False)
    SourceFileName = cString.Trim(SourceFileName)
    DestinationFileName = cString.Trim(DestinationFileName)
    
    If Len(SourceFileName) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyFileName), "SourceFileName")
    If Len(DestinationFileName) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyFileName), "DestinationFileName")
    
    Call Path.VerifyPath(SourceFileName)
    Call Path.VerifyPath(DestinationFileName)
    
    SourceFileName = Path.GetFullPath(cString.TrimEnd(SourceFileName, mDirectorySeparators))
    DestinationFileName = Path.GetFullPath(cString.TrimEnd(DestinationFileName, mDirectorySeparators))
    
    If StrComp(SourceFileName, DestinationFileName, vbTextCompare) = 0 Then Exit Sub
    If API.CopyFile(SourceFileName, DestinationFileName, Not OverWrite) = BOOL_FALSE Then
        Dim e As Long
        e = Err.LastDllError
        If e <> ERROR_FILE_EXISTS Then
            If Not File.Exists(SourceFileName) Then _
                IOError ERROR_FILE_NOT_FOUND, SourceFileName
            If Directory.Exists(DestinationFileName) Then _
                Throw Cor.NewIOException(Environment.GetResourceString(IOException_DirectoryExists))
        End If
        IOError e, SourceFileName
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
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "Path")
    
    Call VBCorLib.Path.VerifyPath(Path)
    If API.DeleteFile(Path) = BOOL_FALSE Then
        Dim Result As Long
        Result = Err.LastDllError
        If Result <> ERROR_FILE_NOT_FOUND Then IOError Result, Path
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
Public Function GetAttributes(ByVal Path As String) As FileAttributes
    Path = cString.Trim(Path)
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "Path")
    
    GetAttributes = API.GetFileAttributes(Path)
    If GetAttributes = INVALID_FILE_ATTRIBUTES Then IOError Err.LastDllError, Path
End Function

''
' Returns the time the file was created in local time.
'
' @param Path The file to retrieve the creation time of.
' @return A cDateTime object representing the creation time.
'
Public Function GetCreationTime(ByVal Path As String) As cDateTime
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    Call FillFileData(Path, Data)
    Set GetCreationTime = cDateTime.FromFileTime(Data.ftCreationTime)
End Function

''
' Returns the time the file was created in UTC time.
'
' @param sPath The file to retrieve the creation time of.
' @return A cDateTime object representing the creation time.
'
Public Function GetCreationTimeUtc(ByVal sPath As String) As cDateTime
    Set GetCreationTimeUtc = GetCreationTime(sPath).ToUniversalTime
End Function

''
' Returns the time the file was accessed in local time.
'
' @param sPath The file to retrieve the last access time for.
' @return A cDateTime object that represents the last time the file was accessed.
'
Public Function GetLastAccessTime(ByVal sPath As String) As cDateTime
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    Call FillFileData(sPath, Data)
    Set GetLastAccessTime = cDateTime.FromFileTime(Data.ftLastAccessTime)
End Function

''
' Returns The last time a file was accessed in UTC time.
'
' @param sPath The file to get the last accessed time for.
' @return A cDateTime object representing the last time the file was accessed.
'
Public Function GetLastAccessTimeUtc(ByVal sPath As String) As cDateTime
    Set GetLastAccessTimeUtc = GetLastAccessTime(sPath).ToUniversalTime
End Function

''
' Returns the last time the file was written to in local time.
'
' @param sPath The file to retrieve the last written time for.
' @return A cDateTime object representing the last time the file was written to.
'
Public Function GetLastWriteTime(ByVal sPath As String) As cDateTime
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    Call FillFileData(sPath, Data)
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
Public Sub Move(ByVal SourceFileName As String, ByVal DestinationFileName As String)
    Call Path.VerifyPath(SourceFileName)
    Call Path.VerifyPath(DestinationFileName)
    
    If Len(SourceFileName) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "SourceFileName")
    If Len(DestinationFileName) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "DestinationFileName")
    
    If Not File.Exists(SourceFileName) Then IOError ERROR_FILE_NOT_FOUND, SourceFileName
    If API.MoveFile(SourceFileName, DestinationFileName) = BOOL_FALSE Then IOError Err.LastDllError, DestinationFileName
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
Public Function OpenFile(ByVal Path As String, ByVal Mode As FileMode, Optional ByVal Access As FileAccess = -1, Optional ByVal Share As FileShare = FileShare.None) As FileStream
    Set OpenFile = Cor.NewFileStream(Path, Mode, Access, Share)
End Function

''
' Opens a file for reading as a FileStream.
'
' @param Path The file to be opened in read mode.
' @return A FileStream used to read from the file.
' @remarks The FileStream object is declared as Cor.NewFileStream(Path, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadShare).
'
Public Function OpenRead(ByVal Path As String) As FileStream
    Set OpenRead = Cor.NewFileStream(Path, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadShare)
End Function

''
' Opens a file for reading through a StreamReader.
'
' @param Path The file to be opened for reading.
' @return A StreamReader used to read from the file with the default encoding (UTF8Encoding).
' @remarks The StreamReader is declared as Cor.NewStreamReader(Path).
'
Public Function OpenText(ByVal Path As String) As StreamReader
    Set OpenText = Cor.NewStreamReader(Path)
End Function

''
' Opens a file for writing as a FileStream.
'
' @param Path The file to be opend for writing to.
' @return A FileStream object used to write to the specified file.
' @remarks The FileStream object is declared as Cor.NewFileStream(Path, FileMode.OpenOrCreate, FileAccess.WriteAccess, FileShare.None).
'
Public Function OpenWrite(ByVal Path As String) As FileStream
    Set OpenWrite = Cor.NewFileStream(Path, FileMode.OpenOrCreate, FileAccess.WriteAccess, FileShare.None)
End Function

''
' Sets the file attributes (ReadOnly, Hidden, ...)
'
' @param sPath The file to set the attributes of.
' @param Attributes The new attributes to set the file to.
' @remarks In order to set a single attribute without unsetting an existing
' attributes, they will first need to be read from the file (GetAttributes), then
' bit manipulated to create the resulting set of desired attributes before setting.
'
Public Sub SetAttributes(ByVal sPath As String, ByVal Attributes As FileAttributes)
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    Call Path.VerifyPath(sPath)
    If API.SetFileAttributes(sPath, Attributes) = BOOL_FALSE Then IOError Err.LastDllError, sPath
End Sub

''
' Sets the time of creation for a file in local time.
'
' @param Path The file to set the time of creation for.
' @param CreationTime A Date or cDateTime object of the new time.
'
Public Sub SetCreationTime(ByVal Path As String, ByVal CreationTime As Variant)
    Call InternalSetFileTime(Path, cDateTime.GetcDateTime(CreationTime).ToFileTime)
End Sub

''
' Sets the time of creation for a file in UTC time.
'
' @param Path The file to set the time of creation for.
' @param CreationTime A Date or cDateTime object of the new time.
'
Public Sub SetCreationTimeUtc(ByVal Path As String, ByVal CreationTime As Variant)
    Call InternalSetFileTime(Path, cDateTime.GetcDateTime(CreationTime).ToFileTimeUtc)
End Sub

''
' Sets the last time the file was accessed in local time.
'
' @param Path The file to set the last access time for.
' @param LastAccessTime A Date or cDateTime object of the new time.
'
Public Sub SetLastAccessTime(ByVal Path As String, ByVal LastAccessTime As Variant)
    Call InternalSetFileTime(Path, , cDateTime.GetcDateTime(LastAccessTime).ToFileTime)
End Sub

''
' Sets the last time the file was accessed in UTC time.
'
' @param Path The file to set the last access time for.
' @param LastAccessTime A Date or cDateTime object of the new time.
'
Public Sub SetLastAccessTimeUtc(ByVal Path As String, ByVal LastAccessTime As Variant)
    Call InternalSetFileTime(Path, , cDateTime.GetcDateTime(LastAccessTime).ToFileTimeUtc)
End Sub

''
' Sets the last time the file was written to in local time.
'
' @param sPath The file to set the last written time for.
' @param LastWriteTime A Date or cDateTime object of the new time.
'
Public Sub SetLastWriteTime(ByVal sPath As String, ByVal LastWriteTime As Variant)
    Call InternalSetFileTime(sPath, , , cDateTime.GetcDateTime(LastWriteTime).ToFileTime)
End Sub

''
' Sets the last time the file was written to in UTC time.
'
' @param sPath The file to set the last written time for.
' @param LastWriteTime A Date or cDateTime object of the new time.
'
Public Sub SetLastWriteTimeUtc(ByVal sPath As String, ByVal LastWriteTime As Variant)
    Call InternalSetFileTime(sPath, , , cDateTime.GetcDateTime(LastWriteTime).ToFileTimeUtc)
End Sub

''
' Reads the entire contents of a file and returns it as a String.
'
' @param Path The path and filename to read in.
' @param Encoding The encoding to be used in reading the file.
' @return A string containing the contents of the file.
'
Public Function ReadAllText(ByVal Path As String, Optional ByVal Encoding As Encoding) As String
    Dim sr As StreamReader
    
    Set sr = Cor.NewStreamReader(Path, Encoding)
    ReadAllText = sr.ReadToEnd
    Call sr.CloseReader
End Function

''
' Reads an entire file into a byte array.
'
' @param Path The path and filename to read into the byte array.
' @return A byte array containing the contents of the file.
'
Public Function ReadAllBytes(ByVal Path As String) As Byte()
    Dim fs      As FileStream
    Dim Ret()   As Byte
    
    Set fs = Cor.NewFileStream(Path, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadWriteShare)
    
    Dim Length As Long
    Length = fs.Length
    If Length > 0 Then
        ReDim Ret(0 To Length - 1)
        Call fs.ReadBlock(Ret, 0, Length)
    Else
        Ret = Cor.NewBytes()
    End If
    Call fs.CloseStream
    ReadAllBytes = Ret
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
    Dim sr As StreamReader
    Set sr = Cor.NewStreamReader(Path, Encoding)
    
    Dim Ret() As String
    ReDim Ret(31)

    Dim i As Long
    Do
        Ret(i) = sr.ReadLine
        If cString.IsNull(Ret(i)) Then Exit Do
        i = i + 1
        If i > UBound(Ret) Then ReDim Preserve Ret(0 To i * 2 - 1)
    Loop
    Call sr.CloseReader

    If i > 0 Then
        ReDim Preserve Ret(0 To i - 1)
    Else
        Ret = Cor.NewStrings
    End If

    ReadAllLines = Ret
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
Public Sub WriteAllText(ByVal Path As String, ByVal Contents As String, Optional ByVal Encoding As Encoding)
    Dim sw As StreamWriter
    
    Set sw = Cor.NewStreamWriter(Path, Encoding)
    Call sw.WriteValue(Contents)
    Call sw.CloseWriter
End Sub

''
' Opens a file, writes all bytes to the file, then closes the file.
'
' @param Path The path and filename of the file to write to.
' @param Bytes The bytes to be written to the file.
' @remarks If the file doesn't exist it will be created. If the file already
' exists, it will be overwritten.
'
Public Sub WriteAllbytes(ByVal Path As String, ByRef Bytes() As Byte)
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException("Bytes", Environment.GetResourceString(ArgumentNull_Array))
    
    Dim fs As FileStream
    Set fs = Cor.NewFileStream(Path, FileMode.Create)
    Call fs.WriteBlock(Bytes, LBound(Bytes), cArray.GetLength(Bytes))
    Call fs.CloseStream
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
Public Sub WriteAllLines(ByVal Path As String, ByRef Contents() As String, Optional ByVal Encoding As Encoding)
    If cArray.IsNull(Contents) Then _
        Throw Cor.NewArgumentNullException("Contents", Environment.GetResourceString(ArgumentNull_Array))
    
    Dim sw As StreamWriter
    Set sw = Cor.NewStreamWriter(Path, Encoding)
    
    Dim i As Long
    For i = LBound(Contents) To UBound(Contents)
        Call sw.WriteLine(Contents(i))
    Next i
    Call sw.CloseWriter
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetFileData(ByVal p As String, ByRef Data As WIN32_FILE_ATTRIBUTE_DATA) As Long
    p = cString.TrimEnd(p, mDirectorySeparators)
    
    Call Path.VerifyPath(p, True)
    If API.GetFileAttributesEx(p, 0, Data) = BOOL_FALSE Then
        GetFileData = GetFileDataFallback(p, Data)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetFileDataFallback(ByRef Path As String, ByRef Data As WIN32_FILE_ATTRIBUTE_DATA) As Long
    Dim FileData As WIN32_FIND_DATA
    
    With API.FindFirstFile(Path, FileData)
        If .IsValid Then
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

Private Sub InternalSetFileTime(ByRef sPath As String, Optional ByVal CreationTime As Variant, Optional ByVal LastAccessTime As Variant, Optional ByVal LastWriteTime As Variant)
    Dim cnt As Long
    Dim lat As Long
    Dim lwt As Long
    
    If Not IsMissing(CreationTime) Then cnt = VarPtr(CreationTime) + 8
    If Not IsMissing(LastAccessTime) Then lat = VarPtr(LastAccessTime) + 8
    If Not IsMissing(LastWriteTime) Then lwt = VarPtr(LastWriteTime) + 8
    
    If Len(sPath) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    Call Path.VerifyPath(sPath)
    
    Dim FileHandle As Long
    FileHandle = API.CreateFile(sPath, FileAccess.WriteAccess, FileShare.None, 0, FileMode.OpenExisting, FILE_ATTRIBUTE_NORMAL, 0)
    If FileHandle = INVALID_HANDLE Then IOError Err.LastDllError, sPath
    If SetFileTime(FileHandle, ByVal cnt, ByVal lat, ByVal lwt) = BOOL_FALSE Then
        Dim e As Long
        e = Err.LastDllError
        Call CloseHandle(FileHandle)
        IOError e, sPath
    End If
    Call CloseHandle(FileHandle)
End Sub

Private Sub FillFileData(ByRef Path As String, ByRef Data As WIN32_FILE_ATTRIBUTE_DATA)
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_EmptyPath), "sPath")
    
    Dim e As Long
    e = GetFileData(Path, Data)
    If e <> NO_ERROR Then IOError e, Path
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mDirectorySeparators = Cor.NewIntegers(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
End Sub
