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
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: ResourceManagerStatic
'

''
' Provides static methods for the <b>ResourceManager</b> object.
'
' @remarks To access these methods, use the <b>ResourceManager.*</b> syntax.
'
'@Folder("CorLib.System.Resources")
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
Public Function CreateFileBasedResourceManager(ByRef ResourceDir As String, Optional ByRef SearchPattern As String = "*") As ResourceManager
    Set CreateFileBasedResourceManager = New ResourceManager
    CreateFileBasedResourceManager.InitFromFolder SearchPattern, ResourceDir
End Function
