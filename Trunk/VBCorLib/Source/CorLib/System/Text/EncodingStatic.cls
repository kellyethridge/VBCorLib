VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingStatic"
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
'    Module: EncodingStatic
'

''
' Provides static methods used to retrieve existing encodings and convert between encodings.
'
' @remarks This class cannot be directly instantiated. To access the methods, use the
' variable name directly.
' <pre>
' Dim en As Encoding
' Set en = Encoding.UTF8
' </pre>
'
' @see Encoding
'
Option Explicit

Private mUTF8               As UTF8Encoding
Private mUTF7               As UTF7Encoding
Private mASCII              As ASCIIEncoding
Private mUnicode            As UnicodeEncoding
Private mBigEndianUnicode   As UnicodeEncoding
Private mEncodings          As New Hashtable
Private mEncodingTable      As New EncodingTable
Private mDefault            As CodePageEncoding


''
' Returns the default ANSI Encoding for the system.
'
' @return Default system encoding.
'
Public Property Get Default() As Encoding
    If mDefault Is Nothing Then Set mDefault = GetEncoding(0)
    Set Default = mDefault
End Property

''
' Returns a UnicodeEncoding object that encodes using BigEndian byte ordering.
'
' @return A single instance of UnicodeEncoding with BigEndian encoding.
' @remarks This encoding is declared as Cor.NewUnicodeEncoding(True, True).
'
Public Property Get BigEndianUnicode() As UnicodeEncoding
    If mBigEndianUnicode Is Nothing Then Set mBigEndianUnicode = Cor.NewUnicodeEncoding(True, True)
    Set BigEndianUnicode = mBigEndianUnicode
End Property

''
' Returns a UnicodeEncoding object that encodes using LittleEndian byte ordering.
'
' @return A single instance of UnicodeEncoding with LittleEndian encoding.
' @remarks This encoding is declared as Cor.NewUnicodeEncoding(False, True).
'
Public Property Get Unicode() As UnicodeEncoding
    If mUnicode Is Nothing Then Set mUnicode = Cor.NewUnicodeEncoding(False, True)
    Set Unicode = mUnicode
End Property

''
' Returns a shared ASCIIEncoding object.
'
' @return A single instance of an ASCIIEncoding.
' @remarks This encoding is declared as New ASCIIEncoding.
'
Public Property Get ASCII() As ASCIIEncoding
    If mASCII Is Nothing Then Set mASCII = New ASCIIEncoding
    Set ASCII = mASCII
End Property

''
' Returns a shared UTF7Encoding object.
'
' @return A single instance of a UTF7Encoding object
' @remarks This encoding is declared as New UTF7Encoding.
'
Public Property Get UTF7() As UTF7Encoding
    If mUTF7 Is Nothing Then Set mUTF7 = New UTF7Encoding
    Set UTF7 = mUTF7
End Property

''
' Returns a shared UTF8Encoding object.
'
' @return A single instance of UTF8Encoding object.
' @remarks This encoding is declared as Cor.NewUTF8Encoding(True)
'
Public Property Get UTF8() As UTF8Encoding
    If mUTF8 Is Nothing Then Set mUTF8 = Cor.NewUTF8Encoding(True)
    Set UTF8 = mUTF8
End Property

''
' Converts a set of bytes from one encoding to another encoding.
'
' @param srcEncoding The encoding to be used to decode the bytes.
' @param dstEncoding The encoding to be used to re-encode the decoded characters.
' @param Bytes The bytes to be converted from one encoding to the other.
' @param Index The starting byte to begin decoding at.
' @param Count The number of bytes to be included in the decoding process.
' @return The encoded bytes.
'
Public Function Convert(ByVal srcEncoding As Encoding, ByVal dstEncoding As Encoding, ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    If srcEncoding Is Nothing Then _
        Throw Cor.NewArgumentNullException("Source Encoding cannot be Nothing.", "srcEncoding")
    If dstEncoding Is Nothing Then _
        Throw Cor.NewArgumentNullException("Destination Encoding cannot be Nothing.", "dstEncoding")
    
    Dim Ch() As Integer
    Ch = srcEncoding.GetChars(Bytes, Index, Count)
    Convert = dstEncoding.GetBytes(Ch)
End Function

''
' Returns an encoding for a specified codepage.
'
' @param CodePageOrName The codepage number or name to encoding with.
' @return An encoding for a specific codepage or name.
'
Public Function GetEncoding(ByRef CodePageOrName As Variant) As Encoding
    Dim cp As Long
    
    Select Case VarType(CodePageOrName)
        Case vbString:                      cp = mEncodingTable.GetCodePage(CodePageOrName)
        Case vbLong, vbInteger, vbByte:     cp = CodePageOrName
        Case Else
            Throw Cor.NewArgumentException("A Code Page or Encoding Name is required.", "CodePageOrName")
    End Select
    
    If cp = 0 Then cp = GetACP
    
    Select Case cp
        Case 1200:      Set GetEncoding = Me.Unicode
        Case 1201:      Set GetEncoding = Me.BigEndianUnicode
        Case 20127:     Set GetEncoding = Me.ASCII
        Case 65000:     Set GetEncoding = Me.UTF7
        Case 65001:     Set GetEncoding = Me.UTF8
        Case Else
            If mEncodings.Contains(cp) Then
                Set GetEncoding = mEncodings(cp)
            Else
                Set GetEncoding = NewCodePageEncoding(cp)
                Call mEncodings.Add(cp, GetEncoding)
            End If
    End Select
End Function

''
' Returns a list of minimal information about each encoding.
'
' @return An array of classes containing a minimal of information.
'
Public Function GetEncodings() As EncodingInfo()
    GetEncodings = EncodingTable.GetEncodings
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get EncodingTable() As EncodingTable
    Set EncodingTable = mEncodingTable
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NewCodePageEncoding(ByVal cp As Long) As CodePageEncoding
    Set NewCodePageEncoding = New CodePageEncoding
    Call NewCodePageEncoding.Init(cp)
End Function

