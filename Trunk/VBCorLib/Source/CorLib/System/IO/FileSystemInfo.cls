VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSystemInfo"
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
'    Module: FileSystemInfo
'

''
' The FileSystemInfo interface is used to access both FileInfo and DirectoryInfo
' objects using the same code. This allows for similar properties and methods
' to be used on both files and directories with a single routine.
'
Option Explicit

''
' Flags used to indicate properties of a file or directory.
'
' @param ReadOnly The file or directory can only be read from.
' @param Hidden The file or directory is to be hidden in normal file listings.
' @param System The file or directory is system object and caution should be used.
' @param DirectoryAttr The object is a directory.
' @param Archive The file or directory is marked for backup.
' @param Device Reserved for future use.
' @param Normal The file is normal and has no other attributes set. This attribute is valid only if used alone.
' @param Temporary The file is temporary. File systems attempt to keep all of the data in memory for quicker access rather than flushing the data back to mass storage. A temporary file should be deleted by the application as soon as it is no longer needed.
' @param SparseFile The file is a sparse file. Sparse files are typically large files whose data are mostly zeros.
' @param ReparsePoint The file contains a reparse point, which is a block of user-defined data associated with a file or a directory.
' @param Compressed The file is compressed.
' @param Offline The file is offline. The data of the file is not immediately available.
' @param NotContentIndexed The file will not be indexed by the operating system's content indexing service.
' @param Encrypted The file or directory is encrypted. For a file, this means that all data in the file is encrypted. For a directory, this means that encryption is the default for newly created files and directories.
'
Public Enum FileAttributes
    ReadOnly = 1
    Hidden = 2
    System = 4
    DirectoryAttr = 16
    Archive = 32
    Device = 64
    Normal = 128
    Temporary = 256
    SparseFile = 512
    ReparsePoint = 1024
    Compressed = 2048
    Offline = 4096
    NotContentIndexed = 8192
    Encrypted = 16384
End Enum

''
' Returns the attributes for the FileSystemInfo entry.
'
' @return The attributes for the FileSystemInfo entry.
'
Public Property Get Attributes() As FileAttributes: End Property

''
' Sets the attributes for the FileSystemInfo entry.
'
' @param RHS The attributes the entry is set to.
'
Public Property Let Attributes(ByVal RHS As FileAttributes): End Property

''
' Returns the time the FileSystemInfo entry was created in local time.
'
' @return A cDateTime object of the creation time.
'
Public Property Get CreationTime() As Variant: End Property

''
' Sets the time the FileSystemInfo entry was created in local time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let CreationTime(ByVal RHS As Variant): End Property

''
' Sets the time the FileSystemInfo entry was created in local time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set CreationTime(ByVal RHS As Variant): End Property

''
' Returns the time the FileSystemInfo entry was last accessed in local time.
'
' @return A cDateTime object of the time.
'
Public Property Get LastAccessTime() As Variant: End Property

''
' Sets the time the FileSystemInfo entry was last accessed in local time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastAccessTime(ByVal RHS As Variant): End Property

''
' Sets the time the FileSystemInfo entry was last accessed in local time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastAccessTime(ByVal RHS As Variant): End Property

''
' Returns the time the FileSystemInfo entry was last written to in local time.
'
' @return A cDateTime object of the time.
'
Public Property Get LastWriteTime() As Variant: End Property

''
' Sets the time the FileSystemInfo entry was last written to in local time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastWriteTime(ByVal RHS As Variant): End Property

''
' Sets the time the FileSystemInfo entry was last written to in local time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastWriteTime(ByVal RHS As Variant): End Property

''
' Returns the time the FileSystemInfo entry was created in UTC time.
'
' @return A cDateTime object of the creation time.
'
Public Property Get CreationTimeUtc() As Variant: End Property

''
' Sets the time the FileSystemInfo entry was created in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let CreationTimeUtc(ByVal RHS As Variant): End Property

''
' Sets the time the FileSystemInfo entry was created in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set CreationTimeUtc(ByVal RHS As Variant): End Property

''
' Returns the time the FileSystemInfo entry was last accessed in UTC time.
'
' @return A cDateTime object of the time.
'
Public Property Get LastAccessTimeUtc() As Variant: End Property

''
' Sets the time the FileSystemInfo entry was last accessed in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastAccessTimeUtc(ByVal RHS As Variant): End Property

''
' Sets the time the FileSystemInfo entry was last accessed in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastAccessTimeUtc(ByVal RHS As Variant): End Property

''
' Returns the time the FileSystemInfo entry was last written to in UTC time.
'
' @return A cDateTime object of the time.
'
Public Property Get LastWriteTimeUtc() As Variant: End Property

''
' Sets the time the FileSystemInfo entry was last written to in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Let LastWriteTimeUtc(ByVal RHS As Variant): End Property

''
' Sets the time the FileSystemInfo entry was last written to in UTC time.
'
' @param RHS A Date or cDateTime object of the new time.
'
Public Property Set LastWriteTimeUtc(ByVal RHS As Variant): End Property

''
' Returns if the FileSystemInfo entry exists.
'
' @return An indication of the entry existing.
'
Public Property Get Exists() As Boolean: End Property

''
' Returns the extension portion of the name.
'
' @return The extension portion of the name.
'
Public Property Get Extension() As String: End Property

''
' Returns the full name of the entry.
'
' @return The full name of the entry.
'
Public Property Get FullName() As String: End Property

''
' Returns the name of the entry.
'
' @return The name of the entry.
Public Property Get Name() As String: End Property

''
' Deletes the entry from disk.
'
Public Sub Delete(): End Sub

''
' Tests for equality with the passed in value.
'
' @param value The value to test for equality against.
' @return An indication of the equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number used to represent this instance.
'
' @return a pseudo-unique number used to represent this instance.
'
Public Function GetHashCode() As Long: End Function

''
' Reloads the entry with the most current information.
'
Public Sub Refresh(): End Sub

''
' Returns a string representation of this instance.
'
' @return a string representation of this instance.
'
Public Function ToString() As String: End Function

