Attribute VB_Name = "StringHelper"
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
' Module: StringHelper
'
Option Explicit

Private mAttachedChars As SafeArray1d


Public Function CharAt(ByRef Source As String, ByVal Index As Long) As Integer
    If LenB(Source) > 0 And Index >= 0 Then
        CharAt = MemWord(StrPtr(Source) + Index * 2)
    End If
End Function

Public Function LastChar(ByRef Source As String) As Integer
    Dim Length As Long
    Length = Len(Source)
    
    If Length > 0 Then
        Dim LastCharIndex As Long
        LastCharIndex = Length - 1
        LastChar = MemWord(StrPtr(Source) + LastCharIndex * 2)
    End If
End Function

Public Function FirstChar(ByRef Source As String) As Integer
    If LenB(Source) > 0 Then
        FirstChar = MemWord(StrPtr(Source))
    End If
End Function

Public Function FirstTwoChars(ByRef Source As String) As DWord
    If LenB(Source) > 0 Then
        FirstTwoChars = MemDWord(StrPtr(Source))
    End If
End Function

Public Function CharCount(ByRef Source As String, ByVal Char As Integer) As Long
    Dim Chars() As Integer
    Dim Length As Long
    
    Length = Len(Source)
    If Length > 0 Then
        Chars = AttachChars(Source)
        
        Dim i As Long
        For i = 0 To Length - 1
            If Chars(i) = Char Then
                CharCount = CharCount + 1
            End If
        Next
        
        DetachChars Chars
    End If
End Function

Private Function AttachChars(ByRef Target As String) As Integer()
    Debug.Assert mAttachedChars.cLocks = 0
    
    mAttachedChars.cbElements = 2
    mAttachedChars.cDims = 1
    mAttachedChars.cElements = Len(Target)
    mAttachedChars.cLocks = 1
    mAttachedChars.pvData = StrPtr(Target)
    SAPtr(AttachChars) = VarPtr(mAttachedChars)
End Function

Private Sub DetachChars(ByRef Chars() As Integer)
    Debug.Assert SAPtr(Chars) = VarPtr(mAttachedChars)
    
    mAttachedChars.cLocks = 0
    SAPtr(Chars) = vbNullPtr
End Sub
