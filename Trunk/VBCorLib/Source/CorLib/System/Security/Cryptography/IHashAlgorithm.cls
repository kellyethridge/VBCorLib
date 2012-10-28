VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IHashAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2006 Kelly Ethridge
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
'    Module: IHashAlgorithm
'

''
' Provides an interface all hash algorithm callbacks must implement.
'
' @remarks The HashAlgorithmBase class uses callbacks to perform hashing functions.
' A callback object must implement this interface to be used.
'
Option Explicit

''
' Requests a callback to hash a set of bytes, maintaining any state needed between calls.
'
' @param Bytes The data to be hashed.
' @param Index The starting index in the data to begin hashing.
' @param Count The number of bytes to be hashed.
' @remarks It is up to the callback to maintain any state needed between calls.
'
Public Sub HashCore(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long): End Sub

''
' Requests the callback to hash any buffered data and return the resulting hash value.
'
' @param Buffer The internally kept buffered data.
' @param BufferCount The number of bytes actually in the buffer.
' @param MessageLength The size of the message in bytes.
' @return The resulting hash value.
'
Public Function HashFinal(ByRef Buffer() As Byte, ByVal BufferCount As Long, ByVal MessageLength As Currency) As Byte(): End Function

''
' Requests that a callback reset itself to prepare for hashing a new value.
'
Public Sub Initialize(): End Sub


