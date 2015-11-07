VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IResourceDecoder"
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
'    Module: IResourceDecoder
'

''
' Provides an interface to allow a <b>ResourceReader</b> to decode
' resource types found within a .RES file.
'
Option Explicit

''
' Returns the type of resource this decoder can decode.
'
' @return Either the resource name or ID.
'
Public Function GetDecodeType() As Variant: End Function

''
' Decodes a raw resource data, returning the number of embedded resources created.
'
' @param Key The key associated with the raw resource data.
' @param Bytes The raw resource data to be decoded.
' @return The number of resources created from the raw data.
' @remarks The decoder needs to maintain the newly decoded resource. Only one
' resource is decoded at a time and then retrieved, so the need to maintain
' multiple decoded resources is unecessary.
'
Public Function Decode(ByVal Key As ResourceKey, ByRef Bytes() As Byte) As Long: End Function

''
' Returns the next resource key and data that was decoded from DecodeResource.
'
' @param ReturnKey The next resource key return to be returned.
' @param ReturnValue The next resource value to be returned.
' @param Returns True if the values were successfully set, otherwise False.
' @remarks Returns the next available decoded resource. Once that resource has
' been returned, it should not be returned again.
' <p>This function will be called as long as it returns True.</p>
'
Public Function GetResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue As Variant) As Boolean: End Function
