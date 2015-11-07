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
'    Module: DriveInfo
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

Private Const PATHTYPE_UNKNOWN  As Long = 0
Private Const PATHTYPE_VOLUME   As Long = 1
Private Const PATHTYPE_SHARE    As Long = 2

Private Const PROP_NAME As String = "Name"

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
    If GetVolumeInformation(mName, vbNullString, 0, ByVal 0&, 0, 0, Buf, Len(Buf)) = BOOL_FALSE Then IOError.WinIOError Err.LastDllError
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
    If GetVolumeInformation(mName, Buf, Len(Buf), ByVal 0&, 0, 0, vbNullString, 0) = BOOL_FALSE Then IOError.WinIOError Err.LastDllError
    VolumeLabel = SysAllocString(StrPtr(Buf))
End Property

''
' Sets the volume name.
'
' @param RHS The new volume name.
'
Public Property Let VolumeLabel(ByVal RHS As String)
    If SetVolumeLabel(mName, RHS) = BOOL_FALSE Then IOError.WinIOError Err.LastDllError
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
    If GetVolumeInformation(mName, vbNullString, 0, SerialNumber, 0, 0, vbNullString, 0) = BOOL_FALSE Then IOError.WinIOError Err.LastDllError
End Property

''
' Returns the amount of free space is available allowed
' for the current account use.
'
' @return The amount of free space available to the current account.
'
Public Property Get AvailableFreeSpace() As Currency
    If GetDiskFreeSpaceEx(mName, AvailableFreeSpace, 0@, 0@) = BOOL_FALSE Then IOError.WinIOError Err.LastDllError
    AvailableFreeSpace = 10000@ * AvailableFreeSpace
End Property

''
' Returns the amount of free space on the drive.
'
' @return The amount of free space.
'
Public Property Get TotalFreeSpace() As Currency
    If GetDiskFreeSpaceEx(mName, 0@, 0@, TotalFreeSpace) = BOOL_FALSE Then IOError.WinIOError Err.LastDllError
    TotalFreeSpace = 10000@ * TotalFreeSpace
End Property

''
' Returns the amount of storage the drive media is capable of holding.
'
' @return The maximum storage amount.
'
Public Property Get TotalSize() As Currency
    If GetDiskFreeSpaceEx(mName, 0@, TotalSize, 0@) = BOOL_FALSE Then IOError.WinIOError Err.LastDllError
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
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Name As String)
    SetName Name
End Sub

Private Sub SetName(ByVal Name As String)
    If Len(Name) = 0 Then _
        Throw Cor.NewArgumentException("Volume name cannot be an empty string.", "Name")
    
    If Len(Name) = 1 Then Name = Name & ":\"
    
    Dim root As String
    Select Case GetPathRootAndType(Name, root)
        Case PATHTYPE_SHARE
            mName = root & "\"
        Case PATHTYPE_VOLUME
            Select Case Asc(root)
                Case corLowerAChar To corLowerZChar, corUpperAChar To corUpperZChar
                    mName = root
                Case Else
                    Throw Cor.NewArgumentException("Drive letters can only be letters a to z, or A to Z.", "Name")
            End Select
        Case Else
            Throw Cor.NewArgumentException("Path must be a root directory ('c:\'), drive letter ('c'), or share name ('\\Server\Share\').", "Name")
    End Select
End Sub

Private Function GetPathRootAndType(ByRef sPath As String, ByRef root As String) As Long
    root = Path.GetPathRoot(sPath)
    If Len(root) < 2 Then Exit Function
    If Left$(root, 2) = "\\" Then
        GetPathRootAndType = PATHTYPE_SHARE
    Else
        GetPathRootAndType = PATHTYPE_VOLUME
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mName = PropBag.ReadProperty(PROP_NAME, "")
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_NAME, mName
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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
