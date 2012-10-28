VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResStringResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: ResStringResource
'

''
' Represents one resource entry in a .RES file.
'
' @remarks A string resource entry contains 16 strings.
' Each string is preceeded by 2 bytes indicating the length
' of the string. The strings are not NULL terminated, and ARE
' 16-bit Unicode characters.
'
' The resource ID is the number of the block that holds the
' 16 strings. The string ID is calculated using the block number
' and the string index into the 16 strings.
'
' StringID = (ResourceID - 1) * 16 + Index
'
' Index is zero based.
'
Option Explicit

Private mStrings(15) As String

''
' Uses the lower 4 bits of the ID as an
' index into the 16 available slots for strings.
'
' @param ID The string ID.
' @param Value The string value.
'
Friend Sub AddString(ByVal ID As Long, ByRef Value As Variant)
    If Len(Value) > 65535 Then _
        Throw Cor.NewArgumentOutOfRangeException("Strings cannot exceed 65535 characters.")
    
    mStrings(ID And &HF) = Value
End Sub

''
' Returns a byte representation of the string array.
'
Friend Function GetEncodedResource() As Byte()
    Dim ret() As Byte
    ReDim ret(0 To CalculateSize - 1)
    
    Dim i As Long
    Dim Index As Long
    For i = 0 To 15
        Index = WriteString(mStrings(i), Index, ret)
    Next i
    
    GetEncodedResource = ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
' Writes string length and string contents to the byte array,
' advancing an index used for the next string.
'
Private Function WriteString(ByRef Value As String, ByVal Index As Long, ByRef Bytes() As Byte) As Long
    ' Access Bytes(Index) and Bytes(Index + 1) as an
    ' Integer and set the value to the string length.
    AsWord(Bytes(Index)) = Len(Value)
    
    ' Move to the next available byte in the array.
    Index = Index + 2
    
    ' Copy the contents of the string to the byte array.
    If LenB(Value) > 0 Then Call CopyMemory(Bytes(Index), ByVal StrPtr(Value), LenB(Value))
    
    ' Return the index of the next available byte in the array.
    WriteString = Index + LenB(Value)
End Function

Private Function CalculateSize() As Long
    Dim i As Long
    For i = 0 To 15
        ' 2 bytes for the length of the string and then the
        ' number of bytes the string takes up.
        CalculateSize = CalculateSize + 2 + LenB(mStrings(i))
    Next i
End Function
