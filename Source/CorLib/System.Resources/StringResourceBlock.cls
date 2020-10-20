VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringResourceBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2018 Kelly Ethridge
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
' Module: StringResourceBlock
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
'@Folder("CorLib.System.Resources")
Option Explicit

Private mStrings(15) As String

''
' Uses the lower 4 bits of the ID as an index into the 16 available slots for strings.
'
' @param ID The string ID.
' @param Value The string value.
'
Friend Sub AddString(ByVal ID As Long, ByRef Value As String)
    mStrings(ID And &HF) = Value
End Sub

''
' Returns a byte representation of the string array.
'
Friend Function GetEncodedResource() As Byte()
    Dim i               As Long
    Dim EncodedData()   As Byte
    Dim Index           As Long
    
    ReDim EncodedData(0 To CalculateSize - 1)
    
    For i = 0 To 15
        Index = WriteString(mStrings(i), Index, EncodedData)
    Next i
    
    GetEncodedResource = EncodedData
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
    If LenB(Value) > 0 Then
        CopyMemory Bytes(Index), ByVal StrPtr(Value), LenB(Value)
    End If
    
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
