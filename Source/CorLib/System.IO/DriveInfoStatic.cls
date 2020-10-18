VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DriveInfoStatic"
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
'    Module: DriveInfoStatic
'

''
' Provides static methods for manipulating DriveInfo objects.
'
' @remark This class cannot be instantiated directly. To access these methods use
' the <b>DriveInfo.*</b> syntax.
'
' @see DriveInfo
' @see Constructors
'
Option Explicit

''
' Represents the type of drives a DriveInfo can be.
'
' @param Unknown The device is of an unknown type.
' @param NoRootDirectory The DriveInfo is no a root directory.
' @param Removable The drive is a removable storage device.
' @param Fixed The device is a fixed disk.
' @param Network The drive is a network drive.
' @param CDRom The drive is a CD-Rom drive.
' @param RamDisk The drive is a RAM disk.
'
Public Enum DriveType
    Unknown = 0
    NoRootDirectory = 1
    Removable = 2
    Fixed = 3
    Network = 4
    CDRom = 5
    RamDisk = 6
End Enum


''
' Returns an array of all the local logical drives.
'
' @return An array of DriveInfo objects.
'
Public Function GetDrives() As DriveInfo()
    Dim Drives() As String
    Drives = Directory.GetLogicalDrives
    
    Dim ret() As DriveInfo
    ReDim ret(0 To UBound(Drives))
    
    Dim i As Long
    For i = 0 To UBound(Drives)
        Set ret(i) = Cor.NewDriveInfo(Drives(i))
    Next i
    
    GetDrives = ret
End Function

