VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectoryInfo"
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
' Module: DirectoryInfo
'

''
' Represents a directory to modify and retrieve information about.
'
' @remarks The DirectoryInfo object does not verify the existence
' of the directory until information has to be read or written to that directory.
'
' @see Constructors
'
Option Explicit
Implements IObject
Implements FileSystemInfo

Private mPath           As String
Private mOriginalPath   As String
Private mIsDirty        As Boolean
Private mFileAttributes As Long
Private mCreationTime   As CorDateTime
Private mLastAccessTime As CorDateTime
Private mLastWriteTime  As CorDateTime


''
' Returns the name of the directory this instance represents.
'
' @param The name of the directory.
'
Public Property Get Name() As String
    Name = Path.GetFileName(mPath)
    If Len(Name) = 0 Then
        Name = mPath
    End If
End Property

''
' Returns the name of the directory, including the full path to the directory.
'
' @param The full path to the directory.
'
Public Property Get FullName() As String
    FullName = mPath
End Property

''
' Returns the directory's attributes (ReadOnly, Hidden, Archive, ...)
'
' @return The directories attributes since the last time Refresh was called.
' @remarks The attributes are not kept current. If the they are changed since
' last checked, then Refresh must be called to retrieve the current settings.
'
Public Property Get Attributes() As FileAttributes
    If mIsDirty Then
        Refresh
    End If
    Attributes = mFileAttributes
End Property

''
' Sets the attributes for the directory.
'
' @param Value The attributes to set for the directory.
' @remarks Setting just one attribute will clear all the other attributes. To
' preserve any attributes, then OR must be used with the attributes to
' be preserved.
'
Public Property Let Attributes(ByVal Value As FileAttributes)
    File.SetAttributes mPath, Value
    MarkDirty
End Property

''
' Returns the creation time for this directory in local time.
'
' @return A CorDateTime object representing the creation time for this directory.
' @remarks The time returned is the UTC time with the timezone offset applied.
' For Pacific Coast, a -8 hours is added to the time before returning.
'
Public Property Get CreationTime() As Variant
    If mIsDirty Then
        Refresh
    End If
    Set CreationTime = mCreationTime
End Property

''
' Sets the creation time for the directory in local time.
'
' @param Value A Date of the new time.
' @remarks The time has the UTC timezone offset applied to it before
' being written to the directory.
'
Public Property Let CreationTime(ByRef Value As Variant)
    Set CreationTime = CorDateTime.GetDateTime(Value)
End Property

''
' Sets the creation time for the directory in local time.
'
' @param Value A CorDateTime object of the new time.
' @remarks The time has the UTC timezone offset applied to it before
' being written to the directory.
'
Public Property Set CreationTime(ByRef Value As Variant)
    If Not TypeOf Value Is CorDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "CreationTime")
    
    Directory.SetCreationTime mPath, Value
    MarkDirty
End Property

''
' Returns the creation time as UTC time.
'
' @return A CorDateTime object of the UTC time.
' @remarks The time returned is the time on the directory without
' any offsets being applied.
'
Public Property Get CreationTimeUtc() As Variant
    If mIsDirty Then
        Refresh
    End If
    Set CreationTimeUtc = mCreationTime.ToUniversalTime
End Property

''
' Sets the creation time as UTC time.
'
' @param Value A Date value as the new time.
' @remarks The time has the UTC timezone offset removed from it before
' being written to the directory.
'
Public Property Let CreationTimeUtc(ByRef Value As Variant)
    Set Me.CreationTimeUtc = CorDateTime.GetDateTime(Value)
End Property

''
' Sets the creation time as UTC time.
'
' @param Value A CorDateTime object of the new time.
' @remarks The time has the UTC timezone offset removed from it before
' being written to the directory.
'
Public Property Set CreationTimeUtc(ByRef Value As Variant)
    If Not TypeOf Value Is CorDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "CreationTime")

    Directory.SetCreationTimeUtc mPath, Value
    MarkDirty
End Property

''
' Returns the last time the directory was accessed in local time.
'
' @return A CorDateTime object of the time.
' @remarks The returned time is the UTC time with the UTC timezone offset applied.
'
Public Property Get LastAccessTime() As Variant
    If mIsDirty Then
        Refresh
    End If
    Set LastAccessTime = mLastAccessTime
End Property

''
' Sets the time the directory was last accessed in local time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Let LastAccessTime(ByRef Value As Variant)
    Set Me.LastAccessTime = CorDateTime.GetDateTime(Value)
End Property

''
' Sets the time the directory was last accessed in local time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Set LastAccessTime(ByRef Value As Variant)
    If Not TypeOf Value Is CorDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastAccessTime")
    
    Directory.SetLastAccessTime mPath, Value
    MarkDirty
End Property

''
' Returns the last time the directory was accessed in UTC time.
'
' @return a CorDateTime object of the time.
'
Public Property Get LastAccessTimeUtc() As Variant
    If mIsDirty Then
        Refresh
    End If
    Set LastAccessTimeUtc = mLastAccessTime.ToUniversalTime
End Property

''
' Sets the last time the directory was accessed in UTC time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Let LastAccessTimeUtc(ByRef Value As Variant)
    Set Me.LastAccessTimeUtc = CorDateTime.GetDateTime(Value)
End Property

''
' Sets the last time the directory was accessed in UTC time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Set LastAccessTimeUtc(ByRef Value As Variant)
    If Not TypeOf Value Is CorDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastAccessTime")
    
    Directory.SetLastAccessTimeUtc mPath, Value
    MarkDirty
End Property

''
' Returns the last time the directory was written to in local time.
'
' @return A CorDateTime object of the time.
'
Public Property Get LastWriteTime() As Variant
    If mIsDirty Then
        Refresh
    End If
    Set LastWriteTime = mLastWriteTime
End Property

''
' Sets the last time the directory was written to in local time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Let LastWriteTime(ByRef Value As Variant)
    Set LastWriteTime = CorDateTime.GetDateTime(Value)
End Property

''
' Sets the last time the directory was written to in local time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Set LastWriteTime(ByRef Value As Variant)
    If Not TypeOf Value Is CorDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastWriteTime")
    
    Directory.SetLastWriteTime mPath, Value
    MarkDirty
End Property

''
' Returns the last time the directory was written to in UTC time.
'
' @return A CorDateTime object of the time.
'
Public Property Get LastWriteTimeUtc() As Variant
    If mIsDirty Then
        Refresh
    End If
    Set LastWriteTimeUtc = mLastWriteTime.ToUniversalTime
End Property

''
' Sets the last time the directory was written to in UTC time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Let LastWriteTimeUtc(ByRef Value As Variant)
    Set LastWriteTimeUtc = CorDateTime.GetDateTime(Value)
End Property

''
' Sets the last time the directory was written to in UTC time.
'
' @param Value A Date or CorDateTime object of the new time.
'
Public Property Set LastWriteTimeUtc(ByRef Value As Variant)
    If Not TypeOf Value Is CorDateTime Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "LastAccessTime")
    
    Directory.SetLastWriteTimeUtc mPath, Value
    MarkDirty
End Property

''
' Reloads all of the attributes and timestamps for the directory.
'
Public Sub Refresh()
    Dim Data As WIN32_FILE_ATTRIBUTE_DATA
    Dim Result As Long
    
    Result = File.GetFileData(mPath, Data)
    If Result <> NO_ERROR Then
        IOError.WinIOError Result, mPath
    End If
    
    With Data
        mFileAttributes = .dwFileAttributes
        Set mCreationTime = CorDateTime.FromFileTime(.ftCreationTime)
        Set mLastAccessTime = CorDateTime.FromFileTime(.ftLastAccessTime)
        Set mLastWriteTime = CorDateTime.FromFileTime(.ftLastWriteTime)
    End With
    
    mIsDirty = False
End Sub

''
' Checks if the directory actually exists.
'
' @return An indication of the directory existing.
'
Public Property Get Exists() As Boolean
    Exists = Directory.Exists(mPath)
End Property

''
' Returns an empty string for directories.
'
' @return An empty string.
'
Public Property Get Extension() As String
' do nothing
End Property

''
' Returns the parent directory as a DirectoryInfo object.
'
' @return A DirectroyInfo object of the parent directory, or Nothing if
' the original directory is a root, such as 'c:\'
Public Property Get Parent() As DirectoryInfo
    Dim ParentDirectoryName As String
    
    ParentDirectoryName = Path.GetDirectoryName(mPath)
    If Len(ParentDirectoryName) > 0 Then
        Set Parent = Cor.NewDirectoryInfo(ParentDirectoryName)
    End If
End Property

''
' Returns a DirectoryInfo object of the root directory of this instance.
'
' @return A DirectoryInfo object of the root.
'
Public Property Get root() As DirectoryInfo
    Set root = Cor.NewDirectoryInfo(Path.GetPathRoot(mPath))
End Property

''
' Creates the path represented by this instance.
'
Public Sub Create()
    Directory.CreateDirectory mPath
End Sub

''
' Creates a directory within this instance, returning the new subdirectory.
'
' @param sPath The new subdirectory to be created.
' @return A DirectoryInfo object of the new subdirectory path.
'
Public Function CreateSubdirectory(ByVal sPath As String) As DirectoryInfo
    Set CreateSubdirectory = Directory.CreateDirectory(Path.Combine(mPath, sPath))
End Function

''
' Deletes this directory if it is empty, or all of its contents and itself
' if Recursive is set to True.
'
' @param Recursive Indicates if all entries in the directory should be deleted.
'
Public Sub Delete(Optional ByVal Recursive As Boolean)
    Directory.Delete mPath, Recursive
End Sub

''
' Returns a list of directories within this directory.
'
' @param SearchPattern A pattern for all directories to match.
' @return A list of all matching directories.
' @remarks If no matching directories were found, then an empty array is returned.
'
Public Function GetDirectories(Optional ByVal SearchPattern As String = "*", Optional ByVal SearchOpt As SearchOption = TopDirectoryOnly) As DirectoryInfo()
    Dim Directories()   As DirectoryInfo
    Dim Paths()         As String
    
    Paths = Directory.GetDirectories(mPath, SearchPattern, SearchOpt)
    
    Dim NumPaths As Long
    NumPaths = CorArray.Length(Paths)
    If NumPaths > 0 Then
        Dim i As Long
        ReDim Directories(0 To NumPaths - 1)
        
        For i = 0 To NumPaths - 1
            Set Directories(i) = Cor.NewDirectoryInfo(Path.Combine(mPath, Paths(i)))
        Next i
    Else
        Directories = CorArray.CreateInstance(vbObject)
    End If
    
    GetDirectories = Directories
End Function

''
' Moves this directory and its contents to the specified path.
'
' @param DestDirName The new path and name of the directory.
'
Public Sub MoveTo(ByVal DestDirName As String)
    Directory.Move mPath, DestDirName
    mPath = Path.GetFullPath(DestDirName)
    mOriginalPath = Path.GetDirectoryName(mPath)
End Sub

''
' Returns a list of file matching the pattern in this directory.
'
' @param SearchPattern The pattern the files must match.
' @return A list of files that matched the pattern.
' @remarks If no files matched the pattern, a zero-length array is returned.
'
Public Function GetFiles(Optional ByVal SearchPattern As String = "*", Optional ByVal SearchOpt As SearchOption = TopDirectoryOnly) As FileInfo()
    Dim Ret()   As FileInfo
    Dim Files() As String
    Files = Directory.GetFiles(mPath, SearchPattern, SearchOpt)
    
    Dim NumFiles As Long
    NumFiles = CorArray.Length(Files)
    If NumFiles > 0 Then
        Dim i As Long
        ReDim Ret(0 To NumFiles - 1)
        For i = 0 To NumFiles - 1
            Set Ret(i) = Cor.NewFileInfo(Path.Combine(mPath, Files(i)))
        Next i
    Else
        Ret = CorArray.CreateInstance(vbObject)
    End If
    
    GetFiles = Ret
End Function

''
' Returns a list of all entries in the directory that match the pattern.
'
' @param SearchPattern The pattern to match against all entries.
' @return A list of all entries that matched the pattern.
' @remarks If no entries matched the pattern, a zero-length array is returned.
'
Public Function GetFileSystemInfos(Optional ByVal SearchPattern As String = "*", Optional ByVal SearchOpt As SearchOption = TopDirectoryOnly) As FileSystemInfo()
    Dim Files() As String
    Files = Directory.GetFiles(mPath, SearchPattern, SearchOpt)
    
    Dim Folders() As String
    Folders = Directory.GetDirectories(mPath, SearchPattern, SearchOpt)
    
    Dim InfoCount As Long
    InfoCount = CorArray.Length(Files) + CorArray.Length(Folders)
    
    Dim Ret() As FileSystemInfo
    If InfoCount > 0 Then
        Dim i As Long
        ReDim Ret(0 To InfoCount - 1)
        For i = 0 To UBound(Files)
            Set Ret(i) = Cor.NewFileInfo(Path.Combine(mPath, Files(i)))
        Next i
        
        Dim j As Long
        j = UBound(Files) + 1
        For i = 0 To UBound(Folders)
            Set Ret(j + i) = Cor.NewDirectoryInfo(Path.Combine(mPath, Folders(i)))
        Next i
    Else
        ' None found, so return an empty array.
        Ret = CorArray.CreateInstance(vbObject)
    End If
    
    GetFileSystemInfos = Ret
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = mOriginalPath
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
Private Sub Class_Initialize()
    MarkDirty
End Sub

Friend Sub Init(ByRef Path As String)
    mOriginalPath = Path
    mPath = Statics.Path.GetFullPath(Path)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MarkDirty()
    mIsDirty = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Init PropBag.ReadProperty("Path")
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Path", mOriginalPath
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
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

