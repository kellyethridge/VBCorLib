VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DriveInfo"
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
' Module: DriveInfo
'

''
' Provides methods to access information about a drive.
'
' @remarks A constructor is used to create a new DriveInfo.
' <pre>
' Set info = NewDriveInfo("c:\")
' </pre>
'
' @see DriveInfoStatic
' @see Constructors
' @see FileInfo
' @see DirectoryInfo
'
Option Explicit
Implements IObject

Private Const PATHTYPE_VOLUME   As Long = 1
Private Const PATHTYPE_SHARE    As Long = 2
Private Const PROP_NAME         As String = "Name"

Private mName As String


''
' Returns the name of the drive.
'
' @return The name of the drive.
'
Public Property Get Name() As String
    Name = mName
End Property

''
' Returns a DirectoryInfo of the root directory for the drive.
'
' @return The root directory of the drive.
'
Public Property Get RootDirectory() As DirectoryInfo
    Set RootDirectory = Cor.NewDirectoryInfo(mName)
End Property

''
' Returns the drive format (NTFS, FAT32, ...)
'
' @return The drive format.
'
Public Property Get DriveFormat() As String
    Dim Buf As String
    
    Buf = String$(32, 0)
    If GetVolumeInformation(mName, vbNullString, 0, ByVal 0&, 0, 0, Buf, Len(Buf)) = BOOL_FALSE Then
        IOError.WinIOError Err.LastDllError
    End If
    
    DriveFormat = SysAllocString(StrPtr(Buf))
End Property

''
' Returns the volume name.
'
' @return The volume name.
'
Public Property Get VolumeLabel() As String
    Dim Buf As String
    
    Buf = String$(32, 0)
    If GetVolumeInformation(mName, Buf, Len(Buf), ByVal 0&, 0, 0, vbNullString, 0) = BOOL_FALSE Then
        IOError.WinIOError Err.LastDllError
    End If
    
    VolumeLabel = SysAllocString(StrPtr(Buf))
End Property

''
' Sets the volume name.
'
' @param Value The new volume name.
'
Public Property Let VolumeLabel(ByRef Value As String)
    If SetVolumeLabel(mName, Value) = BOOL_FALSE Then
        IOError.WinIOError Err.LastDllError
    End If
End Property

''
' Returns the type of drive this instance represents.
'
' @return The type of drive.
'
Public Property Get DriveType() As DriveType
    DriveType = GetDriveType(mName)
End Property

''
' Returns if the drive is ready to be accessed.
'
' @return Whether the drive is ready to be accessed.
'
Public Property Get IsReady() As Boolean
    IsReady = Directory.Exists(mName)
End Property

''
' Returns the serial number for the drive.
'
' @return The serial number.
'
Public Property Get SerialNumber() As Long
    If GetVolumeInformation(mName, vbNullString, 0, SerialNumber, 0, 0, vbNullString, 0) = BOOL_FALSE Then
        IOError.WinIOError Err.LastDllError
    End If
End Property

''
' Returns the amount of free space is available allowed
' for the current account use.
'
' @return The amount of free space available to the current account.
'
Public Property Get AvailableFreeSpace() As Currency
    If GetDiskFreeSpaceEx(mName, AvailableFreeSpace, 0@, 0@) = BOOL_FALSE Then
        IOError.WinIOError Err.LastDllError
    End If
    
    AvailableFreeSpace = 10000@ * AvailableFreeSpace
End Property

''
' Returns the amount of free space on the drive.
'
' @return The amount of free space.
'
Public Property Get TotalFreeSpace() As Currency
    If GetDiskFreeSpaceEx(mName, 0@, 0@, TotalFreeSpace) = BOOL_FALSE Then
        IOError.WinIOError Err.LastDllError
    End If
    
    TotalFreeSpace = 10000@ * TotalFreeSpace
End Property

''
' Returns the amount of storage the drive media is capable of holding.
'
' @return The maximum storage amount.
'
Public Property Get TotalSize() As Currency
    If GetDiskFreeSpaceEx(mName, 0@, TotalSize, 0@) = BOOL_FALSE Then
        IOError.WinIOError Err.LastDllError
    End If
    
    TotalSize = 10000@ * TotalSize
End Property

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = mName
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
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
Friend Sub Init(ByRef Name As String)
    SetName Name
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetName(ByVal Name As String)
    If Len(Name) = 0 Then _
        Throw Cor.NewArgumentException("Volume name cannot be an empty string.", "Name")
    
    If Len(Name) = 1 Then Name = Name & ":\"
    
    Dim Root As String
    Select Case GetPathRootAndType(Name, Root)
        Case PATHTYPE_SHARE
            mName = Root & "\"
        Case PATHTYPE_VOLUME
            Select Case Asc(Root)
                Case vbLowerAChar To vbLowerZChar, vbUpperAChar To vbUpperZChar
                    mName = Root
                Case Else
                    Throw Cor.NewArgumentException("Drive letters can only be letters a to z, or A to Z.", "Name")
            End Select
        Case Else
            Throw Cor.NewArgumentException("Path must be a root directory ('c:\'), drive letter ('c'), or share name ('\\Server\Share\').", "Name")
    End Select
End Sub

Private Function GetPathRootAndType(ByRef sPath As String, ByRef Root As String) As Long
    Root = Path.GetPathRoot(sPath)
    If Len(Root) < 2 Then Exit Function
    If Left$(Root, 2) = "\\" Then
        GetPathRootAndType = PATHTYPE_SHARE
    Else
        GetPathRootAndType = PATHTYPE_VOLUME
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mName = PropBag.ReadProperty(PROP_NAME, "")
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_NAME, mName
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
