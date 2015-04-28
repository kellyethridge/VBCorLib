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


Public Function ContainsNonWhiteSpace(ByRef s As String) As Boolean
    If LenB(s) > 0 Then
        Dim Chars() As Integer
        Chars = AttachChars(s)
        
        Dim i As Long
        For i = 0 To UBound(Chars)
            If Not IsWhiteSpace(Chars(i)) Then
                ContainsNonWhiteSpace = True
                Exit For
            End If
        Next
            
        DetachChars Chars
    End If
End Function

Public Function CharCount(ByRef s As String, ByVal Char As Integer) As Long
    If LenB(s) > 0 Then
        Dim Chars() As Integer
        Chars = AttachChars(s)

        Dim i As Long
        For i = 0 To UBound(Chars)
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
