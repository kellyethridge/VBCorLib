VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: FileInfo
'

''
' Provides methods for manipulating, setting and retrieving information for the specified file.
'
' @remarks The existence of the file is not verified until an action that requires
' accessing the file.
' <pre>
' Dim info As FileInfo
' Set info = NewFileInfo("c:\myfile.txt")
' Debug.Print info.Length
' </pre>
'
' @see Constructors
'
Option Explicit
Implements IObject
Implements FileSystemInfo

Private Const PropOriginalFileName  As String = "OriginalFileName"
Private Const PropFullName          As String = "FullName"
Private Const PropName              As String = "Name"

Private mOriginalFileName   As String
Private mFullName           As String
Private mName               As String
Private mIOResult           As Long
Private mFileAttributes     As FileAttributes
Private mFileSize           As Currency
Private mCreationTimeUtc    As CorDateTime
Private mLastAccessTimeUtc  As CorDateTime
Private mLastWriteTimeUtc   As CorDateTime



''
' Returns the attributes for the file (ReadOnly, Hidden, ...)
'
' @return The attributes of the file.
' @remarks The attributes are represented as bits. Each attributes has a specific bit.
' To test if a specific attribute exists, the attributes need to be ANDed with
' the individual attribute bits.
'
Public Property Get Attributes() As FileAttributes
    If IsDirty Then
        Refresh
    End If
    
    Attributes = mFileAttributes
End Property

''
' Sets the attributes for the file (ReadOnly, Hidden, ...)
'
' @param Value The attributes to set the file to.
' @remarks All attributes are set at once. In order to modify a single attribute,
' the remaining attribute settings need to be included with the modified attribute.
'
Public Property Let Attributes(ByVal Value As FileAttributes)
    Dim Success As Boolean
    Success = CBool(SetFileAttributesW(mFullName, Value))
    
    If Not Success Then
        IOError.WinIOError Err.LastDllError, mFullName
    End If
    
    MarkDirty
End Property

''
' Returns the creation time as a CorDateTime object.
'
' @return The creation time of the file.
'
Public Property Get CreationTime() As Variant
    Set CreationTime = GetCreationTimeUtc.ToLocalTime
End Property

''
' Sets the creation time of the file.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Let CreationTime(ByRef Value As Variant)
    Set CreationTimeUtc = CorDateTime.GetDateTime(Value).ToUniversalTime
End Property

''
' Sets the creation time of the file.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Set CreationTime(ByRef Value As Variant)
    Set CreationTimeUtc = CorDateTime.GetDateTime(Value).ToUniversalTime
End Property

''
' Returns the creation time as a UTC time.
'
' @return The creation time of the file in UTC time.
'
Public Property Get CreationTimeUtc() As Variant
    Set CreationTimeUtc = GetCreationTimeUtc
End Property

''
' Sets the creation time of the file in UTC time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Let CreationTimeUtc(ByRef Value As Variant)
    Set CreationTimeUtc = CorDateTime.GetDateTime(Value)
End Property

''
' Sets the creation time of the file in UTC time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Set CreationTimeUtc(ByRef Value As Variant)
    File.SetCreationTimeUtc mFullName, Value
    MarkDirty
End Property

''
' Returns the last time the file was accessed.
'
' @return A CorDateTime object of the last access time.
'
Public Property Get LastAccessTime() As Variant
    Set LastAccessTime = GetLastAccessTimeUtc.ToLocalTime
End Property

''
' Sets the time the file was last accessed.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Let LastAccessTime(ByRef Value As Variant)
    Set LastAccessTimeUtc = CorDateTime.GetDateTime(Value).ToUniversalTime
End Property

''
' Sets the time the file was last accessed.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Set LastAccessTime(ByRef Value As Variant)
    Set LastAccessTimeUtc = CorDateTime.GetDateTime(Value).ToUniversalTime
End Property

''
' Returns the time the file was last accessed in UTC time.
'
' @return A CorDateTime object of the last accessed UTC time.
'
Public Property Get LastAccessTimeUtc() As Variant
    Set LastAccessTimeUtc = GetLastAccessTimeUtc
End Property

''
' Sets the time the file was last accessed in UTC time.
'
' @param Value A Date or CorDateTime object of the new UTC time.
'
Public Property Let LastAccessTimeUtc(ByRef Value As Variant)
    Set LastAccessTimeUtc = CorDateTime.GetDateTime(Value)
End Property

''
' Sets the time the file was last accessed in UTC time.
'
' @param Value A Date or CorDateTime object of the new UTC time.

Public Property Set LastAccessTimeUtc(ByRef Value As Variant)
    File.SetLastAccessTimeUtc mFullName, Value
    MarkDirty
End Property

''
' Returns the last time the file was written to.
'
' @return A CorDateTime object of the last time the file was written to.
'
Public Property Get LastWriteTime() As Variant
    Set LastWriteTime = GetLastWriteTimeUtc.ToLocalTime
End Property

''
' Sets the last time the file was written to.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Let LastWriteTime(ByRef Value As Variant)
    Set LastWriteTimeUtc = CorDateTime.GetDateTime(Value).ToUniversalTime
End Property

''
' Sets the last time the file was written to.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Set LastWriteTime(ByRef Value As Variant)
    Set LastWriteTimeUtc = CorDateTime.GetDateTime(Value).ToUniversalTime
End Property

''
' Returns the last time the file was accessed in UTC time.
'
' @return A CorDateTime object of the last time the file was written to.
'
Public Property Get LastWriteTimeUtc() As Variant
    Set LastWriteTimeUtc = GetLastWriteTimeUtc
End Property

''
' Sets the last the time file was written to in UTC time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Let LastWriteTimeUtc(ByRef Value As Variant)
    Set LastWriteTimeUtc = CorDateTime.GetDateTime(Value)
End Property

''
' Sets the last time the file was written to in UTC time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Set LastWriteTimeUtc(ByRef Value As Variant)
    File.SetLastWriteTimeUtc mFullName, Value
    MarkDirty
End Property

''
' Returns a DirectoryInfo object of the directory portion of the file.
'
' @return A DirectoryInfo object of the directory.
' @remarks This does not verify the file exists, or the directory.
'
Public Property Get Directory() As DirectoryInfo
    Set Directory = Cor.NewDirectoryInfo(DirectoryName)
End Property

''
' Returns the directory portion of the filename.
'
' @return The directory portion of the filename.
' @remarks This does not very the existence of the file.
'
Public Property Get DirectoryName() As String
    DirectoryName = Path.GetDirectoryName(mFullName)
End Property

''
' Indicates if the file exists on disk.
'
' @return An indication of the existence of the file.
'
Public Property Get Exists() As Boolean
    If IsDirty Then
        Refresh
    End If
    
    Exists = (mIOResult = NO_ERROR)
End Property

''
' Returns the extension portion of the filename.
'
' @return The extension portion of the filename.
'
Public Property Get Extension() As String
    Extension = Path.GetExtension(mFullName)
End Property

''
' Returns a full path to the file, including the filename.
'
' @return A full path to the file.
' @remarks If not root directory portion is found in the
' filename, then the currenct directory is used.
'
Public Property Get FullName() As String
    FullName = mFullName
End Property

''
' Returns the size of the file in bytes.
'
' @return The size of the file.
' @remarks A Currency datatype is used to allow file sizes greater than 2gig.
'
Public Property Get Length() As Currency
    If IsDirty Then
        Refresh
    End If
    
    Length = mFileSize
End Property

''
' Returns the filename portion of the path.
'
' @return The filename and extension.
'
Public Property Get Name() As String
    Name = mName
End Property

''
' Reloads the file attributes and times.
'
' @remarks This will verify that the file exists.
'
Public Sub Refresh()
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    
    mIOResult = File.GetFileData(mFullName, Data)
        
    If mIOResult = NO_ERROR Then
        With Data
            mFileAttributes = .dwFileAttributes
            Set mCreationTimeUtc = CorDateTime.FromFileTimeUtc(.ftCreationTime)
            Set mLastAccessTimeUtc = CorDateTime.FromFileTimeUtc(.ftLastAccessTime)
            Set mLastWriteTimeUtc = CorDateTime.FromFileTimeUtc(.ftLastWriteTime)
        End With
        
        Dim Ptr As Long
        Ptr = VarPtr(mFileSize)
        MemLong(Ptr) = Data.nFileSizeLow
        MemLong(Ptr + 4) = Data.nFileSizeHigh
        mFileSize = 10000@ * mFileSize
    End If
End Sub

''
' Opens a StreamWriter for the file.
'
' @return A Stream writer to the file.
'
Public Function AppendText() As StreamWriter
    Set AppendText = File.AppendText(mFullName)
End Function

''
' Copies this file to a new location, with optional overwrite.
'
' @param DestFileName The new filename this file will be copied as.
' @param OverWrite Permission to overwrite an existing file.
' @remarks The destination must be on the same volume as this file.
'
Public Function CopyTo(ByVal DestFileName As String, Optional ByVal OverWrite As Boolean = False) As FileInfo
    File.Copy mFullName, DestFileName, OverWrite
    Set CopyTo = Cor.NewFileInfo(DestFileName)
End Function

''
' Creates this file and returns a FileStream object to access it.
'
' @return A FileStream object used to access the created file.
' @remarks If the file does not exist it will be created. If it already exists, it will be overwritten.
'
Public Function Create() As FileStream
    Set Create = File.Create(mFullName)
End Function

''
' Returns a StreamWriter to write to this file.
'
' @return The StreamWriter for writing to this file.
' @remarks If the file does not exist it will be created. If it already exists, it will be overwritten.
'
Public Function CreateText() As StreamWriter
    Set CreateText = File.CreateText(mFullName)
End Function

''
' Deletes this file.
'
Public Sub Delete()
    If Not CBool(DeleteFileW(mFullName)) Then
        Dim Result As Long
        Result = Err.LastDllError
        If Result <> ERROR_FILE_NOT_FOUND Then
            IOError.WinIOError Result, mFullName
        End If
    End If
End Sub

''
' Moves this file to a new location.
'
' @param DestFileName The new location and name of the file.
' @remarks The new location must be on the same volume as the source filename.
'
Public Sub MoveTo(ByRef DestFileName As String)
    If Len(DestFileName) = 0 Then _
        Error.Argument Argument_EmptyFileName, "DestFileName"
    
    Dim FullDestFileName As String
    FullDestFileName = Path.GetFullPath(DestFileName)
    
    If Not CBool(MoveFileW(mFullName, FullDestFileName)) Then
        IOError.WinIOError Err.LastDllError
    End If
    
    mOriginalFileName = DestFileName
    mFullName = Path.GetFullPath(DestFileName)
    mName = Path.GetFileName(mFullName)
    MarkDirty
End Sub

''
' Opens the file as a FileStream.
'
' @param Mode The method of opening the file.
' @param Access The read/write operations allowed by this FileStream.
' @param Share The read/write operations allows by other processes.
' @return A FileStream object to this file.
'
Public Function OpenFile(ByVal Mode As FileMode, Optional ByVal Access As FileAccess = ReadWriteAccess, Optional ByVal Share As FileShare = FileShare.None) As FileStream
    Set OpenFile = File.OpenFile(mFullName, Mode, Access, Share)
End Function

''
' Returns a FileStream in ReadOnly mode.
'
' @return A FileStream that has ReadOnly access to the file.
'
Public Function OpenRead() As FileStream
    Set OpenRead = File.OpenRead(mFullName)
End Function

''
' Returns a StreamReader to this file.
'
' @return A StreamReader for reading the file.
' @remarks The StreamReader's default encoding is used.
'
Public Function OpenText() As StreamReader
    Set OpenText = File.OpenText(mFullName)
End Function

''
' Returns a FileStream with write access to the file.
'
' @return A FileStream used to write to the file.
'
Public Function OpenWrite() As FileStream
    Set OpenWrite = File.OpenWrite(mFullName)
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = mFullName
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef FileName As String)
    mOriginalFileName = FileName
    mFullName = Path.GetFullPath(FileName)
    mName = Path.GetFileName(FileName)
End Sub

Private Sub Class_Initialize()
    MarkDirty
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MarkDirty()
    mIOResult = -1
End Sub

Private Property Get IsDirty() As Long
    IsDirty = (mIOResult = -1)
End Property

Private Function GetCreationTimeUtc() As CorDateTime
    If IsDirty Then
        Refresh
    End If
    
    Set GetCreationTimeUtc = mCreationTimeUtc
End Function

Private Function GetLastAccessTimeUtc() As CorDateTime
    If IsDirty Then
        Refresh
    End If
    
    Set GetLastAccessTimeUtc = mLastAccessTimeUtc
End Function

Private Function GetLastWriteTimeUtc() As CorDateTime
    If IsDirty Then
        Refresh
    End If
    
    Set GetLastWriteTimeUtc = mLastWriteTimeUtc
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mOriginalFileName = PropBag.ReadProperty(PropOriginalFileName, "")
    mFullName = PropBag.ReadProperty(PropFullName, "")
    mName = PropBag.ReadProperty(PropName, "")
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PropOriginalFileName, mOriginalFileName
    PropBag.WriteProperty PropFullName, mFullName
    PropBag.WriteProperty PropName, mName
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   FileSystemInfo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get FileSystemInfo_Attributes() As FileAttributes
    FileSystemInfo_Attributes = Attributes
End Property

Private Property Let FileSystemInfo_Attributes(ByVal Value As FileAttributes)
    Attributes = Value
End Property

Private Property Set FileSystemInfo_CreationTime(ByRef Value As Variant)
    Set CreationTime = Value
End Property

Private Property Let FileSystemInfo_CreationTime(ByRef Value As Variant)
    CreationTime = Value
End Property

Private Property Get FileSystemInfo_CreationTime() As Variant
    Set FileSystemInfo_CreationTime = CreationTime
End Property

Private Property Set FileSystemInfo_CreationTimeUtc(ByRef Value As Variant)
    Set CreationTimeUtc = Value
End Property

Private Property Let FileSystemInfo_CreationTimeUtc(ByRef Value As Variant)
    CreationTimeUtc = Value
End Property

Private Property Get FileSystemInfo_CreationTimeUtc() As Variant
    Set FileSystemInfo_CreationTimeUtc = CreationTimeUtc
End Property

Private Sub FileSystemInfo_Delete()
    Delete
End Sub

Private Function FileSystemInfo_Equals(Value As Variant) As Boolean
    FileSystemInfo_Equals = Equals(Value)
End Function

Private Property Get FileSystemInfo_Exists() As Boolean
    FileSystemInfo_Exists = Exists
End Property

Private Property Get FileSystemInfo_Extension() As String
    FileSystemInfo_Extension = Extension
End Property

Private Property Get FileSystemInfo_FullName() As String
    FileSystemInfo_FullName = FullName
End Property

Private Function FileSystemInfo_GetHashCode() As Long
    FileSystemInfo_GetHashCode = GetHashCode
End Function

Private Property Set FileSystemInfo_LastAccessTime(ByRef Value As Variant)
    Set LastAccessTime = Value
End Property

Private Property Get FileSystemInfo_LastAccessTime() As Variant
    Set FileSystemInfo_LastAccessTime = LastAccessTime
End Property

Private Property Let FileSystemInfo_LastAccessTime(ByRef Value As Variant)
    LastAccessTime = Value
End Property

Private Property Set FileSystemInfo_LastAccessTimeUtc(ByRef Value As Variant)
    Set LastAccessTimeUtc = Value
End Property

Private Property Let FileSystemInfo_LastAccessTimeUtc(ByRef Value As Variant)
    LastAccessTimeUtc = Value
End Property

Private Property Get FileSystemInfo_LastAccessTimeUtc() As Variant
    Set FileSystemInfo_LastAccessTimeUtc = LastAccessTimeUtc
End Property

Private Property Set FileSystemInfo_LastWriteTime(ByRef Value As Variant)
    Set LastWriteTime = Value
End Property

Private Property Let FileSystemInfo_LastWriteTime(ByRef Value As Variant)
    LastWriteTime = Value
End Property

Private Property Get FileSystemInfo_LastWriteTime() As Variant
    Set FileSystemInfo_LastWriteTime = LastWriteTime
End Property

Private Property Set FileSystemInfo_LastWriteTimeUtc(ByRef Value As Variant)
    Set LastWriteTimeUtc = Value
End Property

Private Property Let FileSystemInfo_LastWriteTimeUtc(ByRef Value As Variant)
    LastWriteTimeUtc = Value
End Property

Private Property Get FileSystemInfo_LastWriteTimeUtc() As Variant
    Set FileSystemInfo_LastWriteTimeUtc = LastWriteTimeUtc
End Property

Private Property Get FileSystemInfo_Name() As String
    FileSystemInfo_Name = Name
End Property

Private Sub FileSystemInfo_Refresh()
    Refresh
End Sub

Private Function FileSystemInfo_ToString() As String
    FileSystemInfo_ToString = ToString
End Function
