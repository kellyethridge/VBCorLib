VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IResourceWriter"
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
' Module: IResourceWriter
'

''
' Provides an interface for resource writers.
'
' @remarks Implement this interface to allow multiple resource writer
' types to be shared for different resources.
'
Option Explicit

''
' Adds a new resource to the collection of resources to be written.
'
' @param Value The value of the resource.
' @param ResourceName The identifier of the resource value.
' @param ResourceType A number or name of the resource type.
' @param LanguageID A CultureInfo, Culture Name, or LCID.
'
Public Sub AddResource(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant): End Sub

''
' Writes all the resources added by the <a href="IResourceWriter.AddResource.html">AddResource</a> method to the output file or stream.
'
' @remarks <b>Generate</b> does not close the output file or output stream. Instead, consider calling the Close method.
' <p><b>Generate</b> is useful when you want to create a resource file, but you don't want to close the output stream.</p>
' <p><b>Generate</b> can be called only once. After you call <b>Generate</b>, all <a href="IResourceWriter.html">IResourceWriter</a> methods other than <a href="IResourceWriter.CloseWriter.html">CloseWriter</b> will throw an exception.</p>
Public Sub Generate(): End Sub

''
' Closes the underlying resource file or stream, ensuring all the data has been written to the file.
'
' @remarks If necessary, <b>CloseWriter</b> will call the <a href="IResourceWriter.Generate.html">Generate</a> method to output the resources to the underlying file or stream before closing it.
Public Sub CloseWriter(): End Sub

