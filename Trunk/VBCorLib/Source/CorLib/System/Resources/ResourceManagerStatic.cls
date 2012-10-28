VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceManagerStatic"
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
'    Module: ResourceManagerStatic
'

''
' Provides static methods for the <b>ResourceManager</b> object.
'
' @remarks To access these methods, use the <b>ResourceManager.*</b> syntax.
'
Option Explicit

''
' Creates a ResourceManager from a set of .RES files found in the ResourceDir.
'
' @param ResourceDir The folder containing the .RES files.
' @param SearchPattern The pattern used to match .RES files.
' @return A ResourceManager contain resources from a set of .RES files.
' @remarks To load a single specific .RES file use the Cor.NewResourceManager method.
'
' @see Constructors
' @see ResourceManager
'
Public Function CreateFileBasedResourceManager(ByVal ResourceDir As String, Optional ByVal SearchPattern As String = "*") As ResourceManager
    Set CreateFileBasedResourceManager = New ResourceManager
    Call CreateFileBasedResourceManager.InitFromFolder(SearchPattern, ResourceDir)
End Function
