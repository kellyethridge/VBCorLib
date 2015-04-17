VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringReader"
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
'    Module: StringReader
'

''
' Provides methods to read portions of a string.
'
' @see Constructors
' @see TextReader
'
Option Explicit
Implements IObject
Implements TextReader

Private mString     As String
Private mIsClosed   As Boolean
Private mPosition   As Long
Private mBuffer     As WordBuffer
Private mLength     As Long


''
' Closes the reader
'
Public Sub CloseReader()
    mIsClosed = True
End Sub

''
' Returns the next character in the string without consuming
' that character. The character is still the next to be read.
'
' @return The character Unicode value, or -1 if the end of the string was reached.
' @remarks For characters from &h8000 and above, the value will be
' positive still, even though the underlying Integer is negative. This
' is to allow for a -1 to be returned without being confused with an
' actual character in the string.
'
Public Function Peek() As Long
    Call VerifyIsOpen
    If mPosition < mLength Then
        AsWord(Peek) = mBuffer.Data(mPosition)
    Else
        Peek = -1
    End If
End Function

''
' Returns the next character in the string, moving the current
' position to the next character.
'
' @return For characters from &h8000 and above, the value will be
' positive still, even though the underlying Integer is negative. This
' is to allow for a -1 to be returned without being confused with an
' actual character in the string.
'
Public Function Read() As Long
    Call VerifyIsOpen
    If mPosition < mLength Then
        AsWord(Read) = mBuffer.Data(mPosition)
        mPosition = mPosition + 1
    Else
        Read = -1
    End If
End Function

''
' Fills an array with a block of characters from the string.
'
' @param Chars The array to receive the characters read from the string.
' @param Index The starting index in <i>Chars</i> to begin writing.
' @param Count The maximum number of characters to be read from the string.
' @return The actual number of characters read from the string.
' @remarks If not enough characters exist to fill <i>Count</i> characters, then
' fewer characters will be returned than <i>Count</i>.
'
Public Function ReadBlock(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    Call VerifyIsOpen
    
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(Chars), Index, Count)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Chars", Index, "Index", Count, "Count")
    
    If mPosition + Count > mLength Then Count = mLength - mPosition
    Call CopyMemory(Chars(Index), mBuffer.Data(mPosition), Count * 2)
    ReadBlock = Count
End Function

''
' Reads a line terminated by a CarriageReturn or LineFeed from the string.
'
' @return A line from the string or a null string.
' @remarks <p>The returned line does not contain the termination character.
' If the character is a CarriageReturn followed immediately by a LineFeed, then
' the LineFeed character is also consumed.</P>
' <p>If there are no more characters to be read, then Null is returned.</p>
'
Public Function ReadLine() As String
    Call VerifyIsOpen
    
    If mPosition >= mLength Then Exit Function
    
    Dim i As Long
    i = mPosition
    Do While mPosition < mLength
        Select Case mBuffer.Data(mPosition)
            Case vbReturnChar, vbLineFeedChar: Exit Do
        End Select
        mPosition = mPosition + 1
    Loop
    
    ' we ran out of characters looking for Cr and Lf,
    ' so just return the remainder of the string.
    If mPosition >= mLength Then
        ReadLine = Mid$(mString, i + 1)
        Exit Function
    End If
    
    ReadLine = Mid$(mString, i + 1, mPosition - i)
    
    If mBuffer.Data(mPosition) = vbReturnChar Then
        ' consume the found terminator.
        mPosition = mPosition + 1
        ' if a Cr was followed by an Lf then consume the Lf
        If mPosition < mLength Then
            If mBuffer.Data(mPosition) = vbLineFeedChar Then mPosition = mPosition + 1
        End If
    Else
        ' consume the found terminator.
        mPosition = mPosition + 1
    End If
End Function

''
' Returns from the current position in the string to the end.
'
' @return If there are no more characters to be read, then Null is returned.
'
Public Function ReadToEnd() As String
    Call VerifyIsOpen
    If mPosition < mLength Then
        ReadToEnd = Mid$(mString, mPosition + 1)
        mPosition = mLength
    End If
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef s As String, ByVal IsOwner As Boolean)
    If IsOwner Then
        Call Helper.Swap4(ByVal VarPtr(s), mString)
    Else
        mString = s
    End If
    mLength = Len(mString)
    Call InitWordBuffer(mBuffer, StrPtr(mString), mLength)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyIsOpen()
    If mIsClosed Then Throw Cor.NewObjectDisposedException("StringReader", "Cannot read from StringReader when closed.")
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   TextReader Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextReader_CloseReader()
    Call CloseReader
End Sub

Private Function TextReader_Equals(Value As Variant) As Boolean
    TextReader_Equals = Equals(Value)
End Function

Private Function TextReader_GetHashCode() As Long
    TextReader_GetHashCode = GetHashCode
End Function

Private Function TextReader_Peek() As Long
    TextReader_Peek = Peek
End Function

Private Function TextReader_Read() As Long
    TextReader_Read = Read
End Function

Private Function TextReader_ReadBlock(Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    TextReader_ReadBlock = ReadBlock(Chars, Index, Count)
End Function

Private Function TextReader_ReadLine() As String
    TextReader_ReadLine = ReadLine
End Function

Private Function TextReader_ReadToEnd() As String
    TextReader_ReadToEnd = ReadToEnd
End Function

Private Function TextReader_ToString() As String
    TextReader_ToString = ToString
End Function
