VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Encoder"
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
'    Module: Encoder
'

''
' Provides an interface for encoding unicode characters into bytes.
'
Option Explicit

''
' Returns the number of encoded bytes the set of characters would produce.
'
' @param Chars The set of characters that would be used to produce the encoded bytes.
' @param Index The first character in the array to begin the encoding count.
' @param Count The number of characters to be included in the encoding count.
' @param Flush Indicates if the Encoder should maintain state between calls.
' @return The number of bytes that would be produced if the character set was encoded.
' @remarks Setting Flush as false allows multiple bytes that are needed to
Public Function GetByteCount(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByVal Flush As Boolean) As Long: End Function

''
' Returns the encoded bytes generated from a set of unicode characters.
'
' @param Chars The unicode characters to be encoded.
' @param CharIndex The character to begin encoding at.
' @param CharCount The number of characters to be used in the encoding process.
' @param Bytes The encoded set of bytes produced from encoding <i>Chars</i>.
' @param ByteIndex The starting position in the array to place the encoded bytes.
' @param Flush Indicates if the Encoder should maintain state between calls.
' @return The number of bytes actually produced from the encoding.
'
Public Function GetBytes(ByRef Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, ByRef Bytes() As Byte, ByVal ByteIndex As Long, ByVal Flush As Boolean) As Long: End Function

''
' Returns a string representation of this instance.
'
' @return A string that describes this object instance.
'
Public Function ToString() As String: End Function

''
' Used to determine if to objects are of equal value.
'
' @param value An object to be compared to this instance for equality.
' @return An indication of equality.
'
Public Function Equals(ByRef value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number to identify this instance.
'
' @return A pseudo-unique number to identify this instance.
'
Public Function GetHashCode() As Long: End Function
