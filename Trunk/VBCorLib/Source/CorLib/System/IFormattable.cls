VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFormattable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
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
'    Module: IFormattable
'

''
' Provides functionality to format the value of an object into a string representation.
'
Option Explicit

''
' Returns the string representation of the object.
'
' @param Format The format string that describes the requested formatting of the value. <br>- or -<br><b>vbNullString</b> to use the default formatting for the type.
' @param provider A provider that can format the value.<br>- or -<br><b>Nothing</b> to use the default formatter for the current locale.
' @return A string representation of the object with formatting applied.
'
' @see IFormatProvider
'
Public Function ToString(ByVal Format As String, ByVal provider As IFormatProvider) As String: End Function

