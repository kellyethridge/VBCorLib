Attribute VB_Name = "modEncodingHelpers"
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
'    Module: modEncodingHelpers
'
Option Explicit

Private Const BASE64_BYTES As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

' Cache the Base64 encoded character lookup table for quick access.
Public Base64CharToBits()   As Long
Public Base64Bytes()        As Byte


''
' Initialize the encoded character lookup table.

Public Sub InitEncodingHelpers()
    Dim i As Long
    
    ReDim Base64CharToBits(0 To 127)
    For i = 0 To 127
        Base64CharToBits(i) = vbInvalidChar
    Next i
    For i = 0 To 25
        Base64CharToBits(65 + i) = i
        Base64CharToBits(97 + i) = i + 26
    Next i
    For i = 0 To 9
        Base64CharToBits(48 + i) = i + 52
    Next i
    Base64CharToBits(43) = 62
    Base64CharToBits(47) = 63
    
    ReDim Base64Bytes(63)
    For i = 0 To Len(BASE64_BYTES) - 1
        Base64Bytes(i) = Asc(Mid$(BASE64_BYTES, i + 1, 1))
    Next i
End Sub

''
' Attaches either an Integer Array or a String to a Chars Integer
' array, allowing the same access type to both source types.
'
' @param Source Either an Integer Array or a String to attach to.
' @param Chars The array that will be used to access the elements in Source.
' @param CharsSA The SafeArray structure used to represent Chars.
'
Public Sub AttachChars(ByRef Source As Variant, ByRef Chars() As Integer, ByRef CharsSA As SafeArray1d)
    Select Case VarType(Source)
        Case vbString
            CharsSA.cElements = Len(Source)
            CharsSA.pvData = StrPtr(Source)
            CharsSA.cbElements = 2
            CharsSA.cDims = 1
            SAPtr(Chars) = VarPtr(CharsSA)
        
        Case vbIntegerArray
            SAPtr(Chars) = GetArrayPointer(Source)
        
        Case Else
            Throw Cor.NewArgumentException("Chars must be a String or Integer array.", "Chars")
    End Select
End Sub

