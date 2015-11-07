VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IResourceEncoder"
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
'    Module: IResourceEncoder
'

''
' Provides an interface for classes to be used to encode data
' into byte arrays used by the <b>ResourceWriter</b>.
'
' @remarks An encoder needs to hold multiple encoded resources. Once all resources
' have been encoded, each resource will be requested one at a time. The same
' resource should never be returned more than once.
'
Option Explicit

''
' Requests that the encoder release all currently encoded values.
'
Public Sub Reset(): End Sub

''
' Encodes a value into a byte array for later retrieval.
'
' @param Value The value to be encoded into a resource compatible byte array.
' @param ResourceName The identifier for the specific resource. Should be
' converted to a string or numeric value.
' @param ResourceType The type of resource. This can be a string or numeric value.
' If the is value is not supplied, then an attempt to determine the type should
' be made based on the Value itself.
' @param LanguageID The Locale identifier this resource belongs to. If this value
' is not supplied, then the current culture's LCID should be used.
' @return Returns True if the encoder accepted and encoded the value, False otherwise.
' @remarks This function should not throw an exception. If the encoder cannot
' handle the the value for encoding, it should just return False.
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean: End Function

''
' Returns the next available resource from the encoder.
'
' @param ReturnKey This is the unique identifier for the resource to be returned.
' @param ReturnValue A byte array representing the encoded resource value.
' @return Returns True if a resource has been returned, False otherwise.
' @remarks This function should not throw an exception. If there are not more
' resources to be returned, then False should just be returned.
'
Public Function GetEncodedResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue As Variant) As Boolean: End Function

