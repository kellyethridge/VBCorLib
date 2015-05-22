VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitArrayStatic"
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
'    Module: BitArrayStatic
'

''
' Provides creation methods from various types of arrays.
'
' @remarks This class cannot be directly instantiated. In order to
' use the methods in the class, use the <b>BitArray.*</b> syntax.
' <pre>
' Set Bits = BitArray.FromLongs(LongArray)
' </pre>
' @see BitArray
'
Option Explicit

''
' Returns a BitArray object created from an array of Longs. Each bit
' in each long represents an associated bit in the bit array.
'
' @param values An array of Long datatypes.
' @return A BitArray initialized to the number of bits an their
' state in the values array.
'
Public Function FromLongs(ByRef Values() As Long) As BitArray
    Set FromLongs = New BitArray
    FromLongs.InitFromLongs Values
End Function

''
' Returns a BitArray object created from an array of Bytes. Each bit
' in each byte represents an associated bit in the bit array.
'
' @param values An array of Byte datatypes.
' @return A BitArray initialized to the number of bits and their
' state in the values array.
'
Public Function FromBytes(ByRef Values() As Byte) As BitArray
    Set FromBytes = New BitArray
    FromBytes.InitFromBytes Values
End Function

''
' Returns a BitArray object created from an array of Booleans. Each boolean
' represents a single bit in the bit array.
'
' @param values An array of Boolean datatypes.
' @return A BitArray initialized to the number of bits as there are
' boolean elements in the array. Each bit is set to the corrisponding
' boolean value in the array.
'
Public Function FromBooleans(ByRef Values() As Boolean) As BitArray
    Set FromBooleans = New BitArray
    FromBooleans.InitFromBooleans Values
End Function
