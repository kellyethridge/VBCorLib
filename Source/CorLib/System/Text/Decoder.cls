VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Decoder"
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
'    Module: Decoder
'

''
' Provides an interface to be used in converting blocks of bytes into unicode characters.
'
Option Explicit

''
' Returns the number of characters a block of bytes can produce.
'
' @param Bytes The bytes to be used in decoding to characters.
' @param Index The starting index in <i>Bytes</i> to begin decoding from.
' @param Count The number of bytes to be used in the decoding process.
' @return The number of characters that can be produced from <i>Count</i> number of bytes.
'
Public Function GetCharCount(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long) As Long: End Function

''
' Returns the decoded unicode characters from a block of bytes.
'
' @param Bytes The bytes to be used in decoding to characters.
' @param ByteIndex The starting index to begin decoding bytes into characters.
' @param ByteCount The number of bytes to be used in the decoding process.
' @param Chars The output array for the decoded unicode characters.
' @param CharIndex The starting index in the array to begin storing characters.
' @return The number of characters actually decoded.
'
Public Function GetChars(ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, ByRef Chars() As Integer, ByVal CharIndex As Long) As Long: End Function

''
' Returns a string representation of this instance.
'
' @return A string that describes this instance.
'
Public Function ToString() As String: End Function

''
' Returns an indication of equality with the passed in value.
'
' @param value The value to test equality with.
' @param The result of the equality test.
'
Public Function Equals(ByRef value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number that describes this instance.
'
' @return A pseudo-unique number that describes this instance.
'
Public Function GetHashCode() As Long: End Function

