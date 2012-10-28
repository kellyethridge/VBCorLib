VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextReaderStatic"
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
'    Module: TextReaderStatic
'

''
' Provides static functions to be accesses using TextReader.* calling convention.
'
' @see TextReader
'
Option Explicit

Private mNullReader As TextReader

''
' Returns a TextReader that performs no action.
'
' @return The Null TextReader.
' @remarks The reader performs as if reading from past the end
' of the underlying stream. The default behaviour sets the return
' values to the values used to indicate that nothing has been done.<br>
' <b>Read</b> returns -1.<br>
' <b>Peek</b> returns -1.<br>
' <b>ReadBlock</b> returns 0.<br>
' <b>ReadLine</b> returns Null.<br>
' <b>ReadToEnd</b> returns Null.<br>
'
Public Property Get NullReader() As TextReader
    If mNullReader Is Nothing Then Set mNullReader = New NullTextReader
    Set NullReader = mNullReader
End Property

