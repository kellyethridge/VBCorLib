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
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: TextWriter
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
Public Sub WriteValue(ByRef Value As Variant, ParamArray Args() As Variant): End Sub

''
' Writes the string representation of the characters to the stream using both the
' format provider and encoding objects.
'
' @param Chars The characters to be written.
' @param Index The beginning index within the Chars array to begin writing to the stream.
' @param Count The number of characters to write to the stream.
' @remarks Both Index and Count must be excluded or included. Neither argument can be supplied individually.
'
Public Sub WriteChars(ByRef Chars() As Integer, Optional ByRef Index As Variant, Optional ByRef Count As Variant): End Sub

''
' Writes the string representation of the value followed by a NewLine
' to the stream using both the format provider and encoding objects.
'
' @param Value The value to be written.
' @param args Formatting arguments used in formating the Value.
'
Public Sub WriteLine(ParamArray Value() As Variant): End Sub

''
' Writes the string representation of the characters followed by a NewLine to the stream using both the
' format provider and encoding objects.
'
' @param Chars The characters to be written.
' @param Index The beginning index within the Chars array to begin writing to the stream.
' @param Count The number of characters to write to the stream.
' @remarks Both Index and Count must be excluded or included. Neither argument can be supplied individually.
'
Public Sub WriteLineChars(ByRef Chars() As Integer, Optional ByRef Index As Variant, Optional ByRef Count As Variant): End Sub
