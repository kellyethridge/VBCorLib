VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Char"
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
' Module: Char
'
Option Explicit

''
' Indicates if a character is considered white space.
'
' @param Character A string or integer to check if is white space.
' @param Index An index into the string to determine if a specific character is white space.
' This parameter is ignored if a non-string character is passed in.
' @return If the character is considered white space then True is returned, otherwise False is returned.
'
Public Function IsWhiteSpace(ByRef CharOrString As Variant, Optional ByRef Index As Variant) As Boolean
    Select Case VarType(CharOrString)
        Case vbInteger
            IsWhiteSpace = IsWhiteSpaceChar(CharOrString)
        Case vbString
            IsWhiteSpace = IsWhiteSpaceString(CharOrString, Index)
        Case Else
            Error.Argument Argument_InvalidValueType, "String or Integer"
    End Select
End Function

Friend Function IsWhiteSpaceChar(ByVal Character As Integer) As Boolean
    Select Case Character
        Case &H20, &HD, &H9, &HA, &HB, &HC, &H85, &HA0, &H1680, &H180E, _
             &H2000 To &H200A, _
             &H2028, &H2029, &H202F, &H205F, _
             &H3000
            IsWhiteSpaceChar = True
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsWhiteSpaceString(ByRef s As Variant, ByRef Index As Variant) As Boolean
    Dim CharIndex As Long
    CharIndex = OptionalLong(Index, 0)
    
    If CharIndex < 0 Or CharIndex >= Len(s) Then _
        Throw Error.ArgumentOutOfRange("Index")
    
    Dim Ptr As Long
    Ptr = StrPtr(s) + CharIndex * corSizeOfChar
    IsWhiteSpaceString = IsWhiteSpaceChar(MemWord(Ptr))
End Function


