VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextWriter"
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
'    Module: TextWriter
'

''
' Provides an interface used to write text characters to a stream.
'
' @remarks The specific writer will use the supplied or default Encoding scheme
' to convert text characters to byte arrays to be written to the stream.
' @see StreamWriter
' @see StringWriter
'
Option Explicit

''
' Returns the encoding being used by this reader instance.
'
' @return An encoding type used to convert characters to bytes.
' @remarks All characters need to be converted to byte arrays in order
' to be written to a stream. Encodings perform this conversion.
'
Public Property Get Encoding() As Encoding: End Property

''
' Returns the format provider used by this read instance.
'
' @return The format provider used to format values to text.
'
Public Property Get FormatProvider() As IFormatProvider: End Property

''
' Returns the NewLine character used by this reader.
'
' @return Returns the current newline character.
'
Public Property Get NewLine() As String: End Property

''
' Sets the NewLine character used by this writer.
'
' @param RHS The newline characters to be used.
'
Public Property Let NewLine(ByVal RHS As String): End Property

''
' Closes this writer instance.
'
Public Sub CloseWriter(): End Sub

''
' Returns if the value is equal to this instance of <b>TextWriter</b>
'
' @param Value The value to campare equality with.
' @return <b>True</b> if <i>Value</i> equals this class instance, else <b>False</b>.
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Forces any data being cached to be written to the stream immediately.
'
Public Sub Flush(): End Sub
''
' Returns a pseudo-unique number that identifies this object instance.
'
' @return Value that identifies this instance.
'
Public Function GetHashCode() As Long: End Function

''
' Returns a string representation of this instance.
'
' @return A string representation of this instance.
'
Public Function ToString() As String: End Function

''
' Writes the string representation of the value to the stream using both the
' format provider and encoding objects.
'
' @param Value The value to be written.
' @param args Formatting arguments used in formating the Value.
'
Public Sub WriteValue(ByRef Value As Variant, ParamArray args() As Variant): End Sub

''
' Writes the string representation of the value followed by a NewLine
' to the stream using both the format provider and encoding objects.
'
' @param Value The value to be written.
' @param args Formatting arguments used in formating the Value.
'
Public Sub WriteLine(ParamArray Value() As Variant): End Sub

